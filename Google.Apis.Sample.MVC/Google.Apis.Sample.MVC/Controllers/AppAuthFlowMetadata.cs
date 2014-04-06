using System;
using System.Web.Mvc;
using Microsoft.AspNet.Identity;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Mvc;
using Google.Apis.Drive.v2;
using Google.Apis.Util.Store;

namespace Google.Apis.Sample.MVC.Controllers
{
    public class AppAuthFlowMetadata : FlowMetadata
    {
        private static readonly IAuthorizationCodeFlow flow =
            new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer
                {
                    ClientSecrets = new ClientSecrets
                    {
                        ClientId = "187452204195-mtkhnagqb7vdamgg4q0k21do9oiqa93d.apps.googleusercontent.com",
                        ClientSecret = "8wCz9quXgKn1u60zmWkSKpXy"
                    },
                    Scopes = new[] { DriveService.Scope.Drive },
                    // TODO: maybe in a future post I'll demonstrate a new EFDataStore usage.
                    DataStore = new FileDataStore("Google.Apis.Sample.MVCd5")
                });

        public override string GetUserId(Controller controller)
        {
            return controller.User.Identity.GetUserName();
        }

        public override IAuthorizationCodeFlow Flow
        {
            get { return flow; }
        }
    }
}