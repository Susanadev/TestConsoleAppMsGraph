using Azure.Identity;
using Microsoft.Graph;


var scopes = new[] { "User.Read","User.Read.All","Sites.Read.All"};
var interactiveBrowserCredentialOptions = new InteractiveBrowserCredentialOptions
{
  ClientId = "9b649747-f5a3-45a8-aee7-f7549a04b2e8""
};
var tokenCredential = new InteractiveBrowserCredential(interactiveBrowserCredentialOptions);

var graphClient = new GraphServiceClient(tokenCredential, scopes);

var me = await graphClient.Me.GetAsync();
Console.WriteLine($"Hello {me?.DisplayName}!");

//var users = await graphClient.Users.GetAsync();

/* var licenses = await graphClient.SubscribedSkus.GetAsync();

Console.WriteLine($"List of Licenses in tenant:");

foreach(var license in licenses.Value)
{
  Console.WriteLine(license.SkuPartNumber + "  " + license.SkuId + "  " + license.ConsumedUnits);
}

var users = await graphClient.Users.GetAsync((requestConfiguration) =>
{
	requestConfiguration.QueryParameters.Select = new string []{"id", "displayName","mail" };
});

foreach( var user in users.Value)
{
  Console.WriteLine(user.Id + ": " + user.DisplayName + " <" + user.Mail + ">");
}

var mylicenses = await graphClient.Me.LicenseDetails.GetAsync();

Console.WriteLine("mylicenses:");
Console.WriteLine($"I have "+mylicenses.Value.Count+ "assigned");

foreach(var mylicense in mylicenses.Value)
{

  Console.WriteLine(mylicense.SkuPartNumber);
} */

// get user's files & folders in the root
//var oneDriveRoot = await graphClient.Me.DriveGetAsync();
                                  
// display the results
/*foreach (var driveItem in oneDriveRoot.Root)
{
  Console.WriteLine(driveItem.Id + ": " + driveItem.Name);
}*/

//Get files recently used and accessed by the currently signed-in user.
var results = //await graphClient.Me.Insights.Used.GetAsync();
await graphClient.Me.Insights.Used.GetAsync((requestConfiguration) =>
{
	requestConfiguration.QueryParameters.Count = true;
});

Console.WriteLine($"Files I have used or accessed:({results?.OdataCount?.ToString()}):");

foreach (var item in results?.Value!)
{
  //if(item.ResourceVisualization?.MediaType!="text/html")
  //{
    
    Console.WriteLine($"\n Name: {item.ResourceVisualization?.Title} - ({ item.ResourceVisualization?.Type})" );
    Console.WriteLine($" Last Accessed: {item.LastUsed?.LastAccessedDateTime?.LocalDateTime.ToString()}");
    Console.WriteLine($" Last Modified: {item.LastUsed?.LastModifiedDateTime?.LocalDateTime.ToString()}");
    Console.WriteLine($" Link: {item.ResourceReference?.WebUrl}");
    Console.WriteLine($" Preview text: { item.ResourceVisualization?.PreviewText}");
  //}
}

//Get trending files around a specific user (me)
var trendingfiles = await graphClient.Me.Insights.Trending.GetAsync();
                                
Console.WriteLine($"\n Files trending around me: ({trendingfiles?.Value?.Count.ToString()}):");

foreach (var trendingfile in trendingfiles?.Value!)
{
    Console.WriteLine($"\n Name: {trendingfile.ResourceVisualization?.Title } - ({trendingfile.ResourceVisualization?.Type})");
    Console.WriteLine($"  Weight: {trendingfile.Weight}");
    Console.WriteLine($"  Preview Text: { trendingfile.ResourceVisualization?.PreviewText}");
    Console.WriteLine($"  Item Link: { trendingfile.ResourceReference?.WebUrl}");
}

//var sharedFiles = await graphClient.Me.Insights.Shared.GetAsync();

var sharedFiles = await graphClient.Me.Insights.Shared.GetAsync((requestConfiguration) =>
{
	requestConfiguration.QueryParameters.Count = true;
});

Console.WriteLine($"\n Files shared with me:" + sharedFiles?.OdataCount.ToString());

foreach(var sharedfile in sharedFiles?.Value!)
{
  Console.WriteLine($"Name: {sharedfile.ResourceVisualization?.Title } - Type:  {sharedfile.ResourceVisualization?.Type}");
  Console.WriteLine($"Shared by: {sharedfile.LastShared?.SharedBy}");
  Console.WriteLine($"Share date: {sharedfile.LastShared?.SharedDateTime?.LocalDateTime.ToString()}");
}