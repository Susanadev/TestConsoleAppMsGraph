using Azure.Identity;
using Microsoft.Graph;


var scopes = new[] { "User.Read","User.Read.All","Organization.Read.All"};
var interactiveBrowserCredentialOptions = new InteractiveBrowserCredentialOptions
{
  ClientId = "9b649747-f5a3-45a8-aee7-f7549a04b2e8"
};
var tokenCredential = new InteractiveBrowserCredential(interactiveBrowserCredentialOptions);

var graphClient = new GraphServiceClient(tokenCredential, scopes);
int optionselected=0;
string inputvalue=string.Empty;

var me = await graphClient.Me.GetAsync();
Console.WriteLine($"Hello {me?.DisplayName}!");

do{
    Console.WriteLine($"\nEnter an Option or Type in exit to END application");
    Console.WriteLine($"(1) List licenses in the tenant");
    Console.WriteLine($"(2) List licenses assigned to me\n");

    inputvalue = Console.ReadLine()!;
    if(!string.IsNullOrEmpty(inputvalue))
    {
          if(inputvalue.Equals("exit"))
            break;
          else
              int.TryParse(inputvalue, out optionselected);
          
          if(optionselected==1)
          {
              //Get the list of commercial subscriptions that an organization has acquired
              var licenses = await graphClient.SubscribedSkus.GetAsync();

              Console.WriteLine($"\nLICENSES IN TENANT");

              foreach(var license in licenses?.Value!)
              {
                Console.WriteLine($"SkuId: {license.SkuId }  \tConsumed units: {license.ConsumedUnits} \tSkuPart Number: {license.SkuPartNumber}");
              }
          }
          else if (optionselected==2)
          {
              var mylicenses = await graphClient.Me.LicenseDetails.GetAsync();
              Console.WriteLine($"\nYou have {mylicenses!.Value!.Count} licenses assigned");

              foreach(var mylicense in mylicenses.Value)
              {
                  Console.WriteLine($"SkuId: {mylicense.SkuId } \tSkuPart Number: {mylicense.SkuPartNumber}");
              }      
          }
          else 
          {
            Console.WriteLine($"\nInput invalid, please try again!");
          }
    }
    else
    {
        Console.WriteLine($"\nPlease enter a value!!");
    }
} while(!inputvalue.Equals("exit"));