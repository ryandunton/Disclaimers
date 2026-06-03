using Microsoft.AspNetCore.Mvc;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Azure.Security.KeyVault.Secrets;
using Microsoft.Extensions.Caching.Memory;
using Azure.Identity;
using Azure.Core;
using Microsoft.Extensions.Configuration;


// This is a sample Web API controller that returns a list of disclaimers from a SharePoint list.
// A fetch call is used from the office taskpane javascript code to call this controller.  The controller builds a Graph API request to fetch items from a SharePoint list and includes credentials to authenticate the request.
// Requirements:
//  To use this approach, you would need to configure an Azure Web App with managed identity and grant it the necessary permissions to access the SharePoint site.
//  You will need to grant the managed identity permissions to Graph API for Sites.Read.All (or equivalent) to read the SharePoint List.
//  You need to provide the SharePoint site ID and list ID to fetch the list items.
//  The SharePoint list is expected to use a list columns where "Title" is mapped to "Description", "Text" is mapped to "Text", and "Ver" is mapped to "Version".
//  The SharePoint list should also include a "RichText" column for HTML formatted content.
// Notes:
// Version is not used, but provided in the event you need to version responses (you would need to add this code)

namespace SharePointListApi.Controllers
{

    public class ListItemFields
    {
        public required string Description { get; set; }

        public required string Text { get; set; }

        public required string Version { get; set; }

        public string? RichText { get; set; }
    }


    [Route("api/[controller]")]
    [ApiController]
    public class SharePointListController : ControllerBase
    {
        private readonly IHttpClientFactory _clientFactory;
        private readonly SecretClient _secretClient;
        private readonly IMemoryCache _memoryCache;
        private readonly IConfiguration _configuration;
        private readonly IHostEnvironment _environment;

        // Cache keys
        private const string SiteIdCacheKey = "SharePointSiteId";
        private const string ListIdCacheKey = "SharePointListId";
        private const string AccessTokenCacheKey = "SharePointGraphAccessToken";
        
        // For fallback in development - these keys may be needed
        private const string TenantIdCacheKey = "TenantId";
        private const string ClientIdCacheKey = "ClientId";
        private const string ClientSecretCacheKey = "SharePointSecret";
        
        // Cache expiration times (adjust as needed)
        private static readonly TimeSpan SecretCacheTime = TimeSpan.FromHours(12);
        private static readonly TimeSpan TokenCacheTime = TimeSpan.FromMinutes(55); // Token typically lasts 1 hour

        public SharePointListController(
            IHttpClientFactory clientFactory, 
            SecretClient secretClient, 
            IMemoryCache memoryCache,
            IConfiguration configuration,
            IHostEnvironment environment)
        {
            _clientFactory = clientFactory;
            _secretClient = secretClient;
            _memoryCache = memoryCache;
            _configuration = configuration;
            _environment = environment;
        }

        [HttpGet]
        public async Task<IActionResult> Get()
        {
            try
            {
                // Get site ID and list ID from Key Vault
                var siteId = await GetSecretFromCacheOrVault(SiteIdCacheKey); 
                var listId = await GetSecretFromCacheOrVault(ListIdCacheKey);

                if (string.IsNullOrEmpty(siteId) || string.IsNullOrEmpty(listId))
                {
                    return StatusCode(500, new { error = "SharePoint site ID or list ID is missing in configuration." });
                }

                // Get an access token
                string token;
                
                try
                {
                    // Get an access token using managed identity, or fall back to client credentials if in development
                    token = await GetCachedAccessToken();
                
                    // Retrieve items from the SharePoint list
                    var listItems = await GetSharePointListItems(token, siteId, listId);
                    return Ok(listItems);
                }
                catch (Exception authEx)
                {
                    // Provide more detailed error message for authentication issues
                    var errorDetail = _environment.IsDevelopment() 
                        ? $"Authentication error: {authEx.Message}" 
                        : "Authentication error occurred. Check application logs for details.";
                    
                    return StatusCode(500, new { error = errorDetail });
                }
            }
            catch (Exception ex)
            {
                // Provide a proper JSON response for errors
                var errorMessage = _environment.IsDevelopment()
                    ? $"An error occurred while fetching the list items: {ex.Message}"
                    : "An error occurred while processing your request.";
                
                return StatusCode(500, new { error = errorMessage });
            }
        }

        /// <summary>
        /// Gets a secret from cache if available, otherwise retrieves it from Key Vault and caches it
        /// </summary>
        private async Task<string> GetSecretFromCacheOrVault(string secretName)
        {
            // Try to get from cache first
            if (_memoryCache.TryGetValue(secretName, out string cachedSecret))
            {
                return cachedSecret;
            }

            try
            {
                // Not in cache, get from Key Vault
                var secret = await _secretClient.GetSecretAsync(secretName);
                var secretValue = secret.Value.Value;

                // Cache the secret with expiration
                _memoryCache.Set(secretName, secretValue, SecretCacheTime);
                
                return secretValue;
            }
            catch (Exception ex)
            {
                // Log the exception
                Console.WriteLine($"Error retrieving secret {secretName}: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// Gets an access token from cache if available, otherwise retrieves a new one and caches it
        /// </summary>
        private async Task<string> GetCachedAccessToken()
        {
            // Try to get from cache first
            if (_memoryCache.TryGetValue(AccessTokenCacheKey, out string cachedToken))
            {
                return cachedToken;
            }

            // Not in cache, get a new token
            string token;
            
            // Try managed identity first
            try
            {
                token = await GetAccessTokenWithManagedIdentity();
            }
            catch (Exception ex) when (_environment.IsDevelopment())
            {
                // In development, fall back to client credentials if managed identity fails
                Console.WriteLine($"Managed identity authentication failed: {ex.Message}");
                Console.WriteLine("Attempting fallback to client credentials for development...");
                
                var tenantId = await GetSecretFromCacheOrVault(TenantIdCacheKey);
                var clientId = await GetSecretFromCacheOrVault(ClientIdCacheKey);
                var clientSecret = await GetSecretFromCacheOrVault(ClientSecretCacheKey);
                
                if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret))
                {
                    throw new Exception("Client credentials not available for fallback authentication.");
                }
                
                token = await GetAccessTokenWithClientCredentials(tenantId, clientId, clientSecret);
            }
            
            // Cache the token with expiration
            _memoryCache.Set(AccessTokenCacheKey, token, TokenCacheTime);
            
            return token;
        }

        /// <summary>
        /// Gets an access token using managed identity
        /// </summary>
        private async Task<string> GetAccessTokenWithManagedIdentity()
        {
            try
            {
                // Use DefaultAzureCredential which supports managed identity in Azure
                var credential = new DefaultAzureCredential(new DefaultAzureCredentialOptions 
                { 
                    AuthorityHost = AzureAuthorityHosts.AzureGovernment 
                });
                
                // Get token for Microsoft Graph
                var tokenRequestContext = new TokenRequestContext(
                    new[] { "https://graph.microsoft.us/.default" });
                
                var tokenResult = await credential.GetTokenAsync(tokenRequestContext);
                
                return tokenResult.Token;
            }
            catch (Exception ex)
            {
                // Log the exception and rethrow
                throw new Exception($"Failed to acquire token with managed identity: {ex.Message}", ex);
            }
        }
        
        /// <summary>
        /// Gets an access token using client credentials (fallback for development only)
        /// </summary>
        private async Task<string> GetAccessTokenWithClientCredentials(string tenantId, string clientId, string clientSecret)
        {
            try
            {
                var client = _clientFactory.CreateClient();

                var request = new HttpRequestMessage(HttpMethod.Post, $"https://login.microsoftonline.us/{tenantId}/oauth2/v2.0/token");
                request.Content = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("client_id", clientId),
                    new KeyValuePair<string, string>("scope", "https://graph.microsoft.us/.default"),
                    new KeyValuePair<string, string>("client_secret", clientSecret),
                    new KeyValuePair<string, string>("grant_type", "client_credentials"),
                });

                var response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var responseStream = await response.Content.ReadAsStringAsync();
                var responseObject = JsonConvert.DeserializeObject<dynamic>(responseStream);

                return responseObject.access_token;
            }
            catch (Exception ex)
            {
                // Log the exception and rethrow
                throw new Exception($"Failed to acquire token with client credentials: {ex.Message}", ex);
            }
        }

        private async Task<IEnumerable<ListItemFields>> GetSharePointListItems(string accessToken, string siteId, string listId)
        {
            try
            {
                // Set to true if you want to prioritize the order of disclaimers in the list (must define a Priorty field in SharePoint list you configure)
                bool bPrioritizeList = false;

                var client = _clientFactory.CreateClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                var requestUrl = $"https://graph.microsoft.us/beta/sites/{siteId}/lists/{listId}/items?expand=fields(select=Title,Text,Ver,RichText)";

                if (bPrioritizeList)
                {
                    // If you want to prioritize the order of disclaimers, define a "Priority" field as Number in your list and index it
                    // You can use this field to order the disclaimer items in the graph query 
                    // If you are using a Priority field to order disclaimer items, this may fail until you index the field in SharePoint
                    client.DefaultRequestHeaders.Add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
                    requestUrl = $"https://graph.microsoft.us/beta/sites/{siteId}/lists/{listId}/items?expand=fields(select=Title,Text,Ver,RichText)&$orderby=fields/Priority";
                }

                var response = await client.GetAsync(requestUrl);
                response.EnsureSuccessStatusCode();

                var responseStream = await response.Content.ReadAsStringAsync();
                var responseObject = JsonConvert.DeserializeObject<dynamic>(responseStream);

                // Initialize a list to hold the simplified list item fields
                var items = new List<ListItemFields>();

                // Iterate through each item in the response and extract the fields
                // If the columns of your SharePoint list are different, you will need to adjust the fields accordingly
                foreach (var item in responseObject.value)
                {
                    var fields = new ListItemFields
                    {
                        // Map the fields from your response to the properties, assuming `item.fields` contains them
                        Text = item.fields.Text,
                        Description = item.fields.Title,
                        Version = item.fields.Ver,
                        RichText = item.fields.RichText
                    };

                    items.Add(fields);
                }

                // Return the items
                return items;
            }
            catch(HttpRequestException httpEx)
            {
                // Log the exception details or handle them as needed
                throw new Exception($"An error occurred while fetching the list items: {httpEx.Message}", httpEx);
            }
            catch (Exception ex)
            {
                // Handle other exceptions
                throw new Exception($"An unexpected error occurred: {ex.Message}", ex);
            }   
        }
    }
}
