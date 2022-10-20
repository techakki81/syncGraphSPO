
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Text;
using System.Threading.Tasks;
using System.Security;
using System.IO;
using System.Drawing;
using System.Xml.Serialization;
using System.Reflection;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Configuration;
using Azure.Identity;
using Microsoft.Graph;

namespace Contoso.Core.ProfilePictureUploader
{
    class Program
    {

        //sizes for profile pictures
        const int _smallThumbWidth = 48;
        const int _mediumThumbWidth = 72;
        const int _largeThumbWidth = 200;



        static string tenantId = "XXX";
        static string clientId = "XXX";
        static string clientSecret = "XXX";

        static UPSvc.UserProfileService _userProfileService;
        static ClientContext _clientContext;
        const string _sPOProfilePrefix = "i:0#.f|membership|";
        const string _profileSiteUrl = "https://XXX-admin.sharepoint.com";
        const string _mySiteUrl = "https://XXX-my.sharepoint.com";
        const string targetLibraryPath = "/User Photos/Profile Pictures";

        enum LogLevel { Information, Warning, Error };
        static Configuration _appConfig;
        static void Main(string[] args)
        {
            SetSPProfPicFromGraph();
        }

        public static async Task<List<Microsoft.Graph.User>> getUsers(GraphServiceClient graphClient)
        {
            List<Microsoft.Graph.User> usersList = new List<Microsoft.Graph.User>();

           
            var users = await graphClient.Users
                .Delta()
                .Request()
                .GetAsync();

            usersList.AddRange(users.CurrentPage);
            try
            {
                while (users.NextPageRequest != null)
                {
                    users = await users.NextPageRequest.GetAsync();
                    usersList.AddRange(users.CurrentPage);
                }
            }
            catch (Exception e)
            {
                // log it
            }
            return usersList;
        }

        public static void SetSPProfPicFromGraph()
        {

            InitializeWebService();

           
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            
            var clientSecretCredential = new ClientSecretCredential(
                tenantId, clientId, clientSecret, options);

            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            var allUsers = getUsers(graphClient).Result;
            foreach (var s in allUsers)
            {
                try
                {
                    Console.WriteLine(allUsers.Count);
                    Console.WriteLine(s.UserPrincipalName);
                    Console.WriteLine(s.UserType);

                   
                    var requestUserPhoto = graphClient.Users[s.UserPrincipalName].Photo.Request();
                    var resultsUserPhoto = requestUserPhoto.GetAsync().Result;

                    // get user photo content
                    var requestUserPhotoFile = graphClient.Users[s.UserPrincipalName].Photos["64x64"].Content.Request();
                    var resultsUserPhotoFile = requestUserPhotoFile.GetAsync().Result;

                    string sPoUserProfileName = s.UserPrincipalName;
                    //create SP naming convetion for image file
                    string newImageNamePrefix = sPoUserProfileName.Replace("@", "_").Replace(".", "_");
                    //upload source image to SPO (might do some resize work, and multiple image upload depending on config file)
                    string spoImageUrl = UploadImageToSpo(newImageNamePrefix, resultsUserPhotoFile);
                    if (spoImageUrl.Length > 0)//if upload worked
                    {
                        string[] profilePropertyNamesToSet = new string[] { "PictureURL", "SPS-PicturePlaceholderState", "SPS-PictureTimestamp" };
                        string[] profilePropertyValuesToSet = new string[] { spoImageUrl, "0", "63605901091" };
                        //set these 2 required properties for user profile i.e path to image uploaded, and pictureplaceholder state
                        SetMultipleProfileProperties(_sPOProfilePrefix + sPoUserProfileName, profilePropertyNamesToSet, profilePropertyValuesToSet);                       
                    }                  
               
                }
                catch (Exception ex)
                {
                    // log 
                    string ds = ex.InnerException.Message;

                    continue;
                }
            }
        }



        

        
        /// <summary>
        /// Upload picture stream to SPO My Site hoste (SkyDrive Pro Host) site collection user photos library.
        /// </summary>
        /// <param name="PictureName"></param>
        /// <param name="ProfilePicture"></param>
        /// <returns>URL to uploaded picture</returns>
        static string UploadImageToSpo(string PictureName, Stream ProfilePicture)
        {
            try
            {
                string spPhotoPathTempate = string.Concat(targetLibraryPath.TrimEnd('/'), "/{0}_{1}Thumb.jpg"); //path template to photo lib in My Site Host
                string spImageUrl = string.Empty;

                //create SPO Client context to My Site Host
                ClientContext mySiteclientContext = new ClientContext(_mySiteUrl);
                SecureString securePassword = GetSecurePassword(Constants._sPoAuthPasword);
                //provide auth crendentials using O365 auth
                mySiteclientContext.Credentials = new SharePointOnlineCredentials(Constants._sPoAuthUserName, securePassword);

                {
                    LogMessage("Uploading threes image to SPO, with resizing", LogLevel.Information);
                    //create 3 images based on recommended sizes for SPO
                    //create small size,                   
                    using (Stream smallThumb = ResizeImageSmall(ProfilePicture, _smallThumbWidth))
                    {
                        if (smallThumb != null)
                        {
                            spImageUrl = string.Format(spPhotoPathTempate, PictureName, "S");
                            LogMessage("Uploading small image to " + spImageUrl, LogLevel.Information);
                            Microsoft.SharePoint.Client.File.SaveBinaryDirect(mySiteclientContext, spImageUrl, smallThumb, true);
                        }
                    }

                    //create medium size
                    using (Stream mediumThumb = ResizeImageSmall(ProfilePicture, _mediumThumbWidth))
                    {
                        if (mediumThumb != null)
                        {
                            spImageUrl = string.Format(spPhotoPathTempate, PictureName, "M");
                            LogMessage("Uploading medium image to " + spImageUrl, LogLevel.Information);
                            Microsoft.SharePoint.Client.File.SaveBinaryDirect(mySiteclientContext, spImageUrl, mediumThumb, true);

                        }
                    }

                    //create large size image, shown on SkyDrive Pro main page for user
                    using (Stream largeThumb = ResizeImageLarge(ProfilePicture, _largeThumbWidth))
                    {
                        if (largeThumb != null)
                        {

                            spImageUrl = string.Format(spPhotoPathTempate, PictureName, "L");
                            LogMessage("Uploading large image to " + spImageUrl, LogLevel.Information);
                            Microsoft.SharePoint.Client.File.SaveBinaryDirect(mySiteclientContext, spImageUrl, largeThumb, true);

                        }
                    }


                }
                //return medium sized URL, as this is the one that should be set in the user profile
                return _mySiteUrl + string.Format(spPhotoPathTempate, PictureName, "M");

            }
            catch (Exception ex)
            {
                LogMessage("User Error: Failed to upload thumbnail picture to SPO for " + PictureName + " " + ex.Message, LogLevel.Error);
                return string.Empty;
            }

        }


        /// <summary>
        /// Resize image stream to width passed into function. Will use source image dimension to scale image correctly
        /// </summary>
        /// <param name="OriginalImage"></param>
        /// <param name="NewWidth">New image size width in pixels</param>
        /// <returns></returns>
        static Stream ResizeImageSmall(Stream OriginalImage, int NewWidth)
        {

         
            try
            {
                OriginalImage.Seek(0, SeekOrigin.Begin);
                System.Drawing.Image originalImage = System.Drawing.Image.FromStream(OriginalImage, true, true);
                if (originalImage.Width == NewWidth) //if sourceimage is same as destination, no point resizing, as it loses quality
                {
                    OriginalImage.Seek(0, SeekOrigin.Begin);
                    originalImage.Dispose();
                    return OriginalImage; //return same image that was passed in
                }
                else
                {
                    System.Drawing.Image resizedImage = originalImage.GetThumbnailImage(NewWidth, (NewWidth * originalImage.Height) / originalImage.Width, null, IntPtr.Zero);
                    MemoryStream memStream = new MemoryStream();
                    resizedImage.Save(memStream, ImageFormat.Jpeg);
                    resizedImage.Dispose();
                    originalImage.Dispose();
                    memStream.Seek(0, SeekOrigin.Begin);
                    return memStream;
                }


            }
            catch (Exception ex)
            {
                LogMessage("User Error: cannot create resized image to new width of " + NewWidth.ToString() + ex.Message, LogLevel.Error);
                return null;
            }
        }


        /// <summary>
        /// Delivers better quality image for scaling large thumbs e.g. 200px in width
        /// </summary>
        /// <param name="OriginalImage"></param>
        /// <param name="NewWidth"></param>
        /// <returns></returns>
        static Stream ResizeImageLarge(Stream OriginalImage, int NewWidth)
        {
            OriginalImage.Seek(0, SeekOrigin.Begin);
            System.Drawing.Image originalImage = System.Drawing.Image.FromStream(OriginalImage, true, true);
            int newHeight = (NewWidth * originalImage.Height) / originalImage.Width;

            Bitmap newImage = new Bitmap(NewWidth, newHeight);

            using (Graphics gr = Graphics.FromImage(newImage))
            {
                gr.SmoothingMode = SmoothingMode.HighQuality;
                gr.InterpolationMode = InterpolationMode.HighQualityBicubic;
                gr.PixelOffsetMode = PixelOffsetMode.HighQuality;
                gr.DrawImage(originalImage, new Rectangle(0, 0, NewWidth, newHeight)); //copy to new bitmap
            }


            MemoryStream memStream = new MemoryStream();
            newImage.Save(memStream, ImageFormat.Jpeg);
            originalImage.Dispose();
            memStream.Seek(0, SeekOrigin.Begin);
            return memStream;


        }

        /// <summary>
        /// Help funtion to log messages to the console window, and to a text file. Log level is currently not used other than for display colors in console
        /// </summary>
        /// <param name="Message"></param>
        /// <param name="Level"></param>
        static void LogMessage(string Message, LogLevel Level)
        {
            //maybe write to log where image failed to upload or profile picture
            switch (Level)
            {
                case LogLevel.Error: Console.ForegroundColor = ConsoleColor.Red; break;
                case LogLevel.Warning: Console.ForegroundColor = ConsoleColor.Green; break;
                case LogLevel.Information: Console.ForegroundColor = ConsoleColor.White; break;

            }

            Console.WriteLine(Message);
            Console.ResetColor();

            try
            {

                if (_appConfig != null)
                    if (_appConfig.LogFile.EnableLogging) //check if logging is enabled in configuration file
                    {
                        System.IO.File.AppendAllText(_appConfig.LogFile.Path, Environment.NewLine + DateTime.Now + " : " + Message);
                    }

            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error writing to log file. " + ex.Message);
                Console.ResetColor();
            }
        }


        /// <summary>
        /// Use the properties in the configuration file and set these against the user profile
        /// </summary>
        /// <param name="UserName"></param>
        static void SetAdditionalProfileProperties(string UserName)
        {
            if (_appConfig.AdditionalProfileProperties.Properties == null) //if properties has been left out of config file
                return;

            int propsCount = _appConfig.AdditionalProfileProperties.Properties.Count();
            if (propsCount > 0)
            {
                string[] profilePropertyNamesToSet = new string[propsCount];
                string[] profilePropertyValuesToSet = new string[propsCount];
                //loop through each property in config
                for (int i = 0; i < propsCount; i++)
                {
                    profilePropertyNamesToSet[i] = _appConfig.AdditionalProfileProperties.Properties[i].Name;
                    profilePropertyValuesToSet[i] = _appConfig.AdditionalProfileProperties.Properties[i].Value;
                }

                //set all props in a single call
                SetMultipleProfileProperties(UserName, profilePropertyNamesToSet, profilePropertyValuesToSet);

            }
        }


        /// <summary>
        /// Use this function if you want to set a single property in the user profile store
        /// </summary>
        /// <param name="UserName"></param>
        /// <param name="PropertyName"></param>
        /// <param name="PropertyValue"></param>
        static void SetSingleProfileProperty(string UserName, string PropertyName, string PropertyValue)
        {

            try
            {


                UPSvc.PropertyData[] data = new UPSvc.PropertyData[1];
                data[0] = new UPSvc.PropertyData();
                data[0].Name = PropertyName;
                data[0].IsValueChanged = true;
                data[0].Values = new UPSvc.ValueData[1];
                data[0].Values[0] = new UPSvc.ValueData();
                data[0].Values[0].Value = PropertyValue;
                _userProfileService.ModifyUserPropertyByAccountName(UserName, data);


            }
            catch (Exception ex)
            {
                LogMessage("Exception trying to update profile property " + PropertyName + " for user " + UserName + "\n" + ex.Message, LogLevel.Error);
            }

        }

        /// <summary>
        /// Use this function is in a single call to SPO you want to set multiple profile properties. 1st item in propertyname array is associated to first item in propertvalue array etc
        /// </summary>
        /// <param name="UserName"></param>
        /// <param name="PropertyName"></param>
        /// <param name="PropertyValue"></param>
        static void SetMultipleProfileProperties(string UserName, string[] PropertyName, string[] PropertyValue)
        {

            LogMessage("Setting multiple SPO user profile properties for " + UserName, LogLevel.Information);

            try
            {
                int arrayCount = PropertyName.Count();

                UPSvc.PropertyData[] data = new UPSvc.PropertyData[arrayCount];
                for (int x = 0; x < arrayCount; x++)
                {
                    data[x] = new UPSvc.PropertyData();
                    data[x].Name = PropertyName[x];
                    data[x].IsValueChanged = true;
                    data[x].Values = new UPSvc.ValueData[1];
                    data[x].Values[0] = new UPSvc.ValueData();
                    data[x].Values[0].Value = PropertyValue[x];
                }

                _userProfileService.ModifyUserPropertyByAccountName(UserName, data);
                LogMessage("Finished setting multiple SPO user profile properties for " + UserName, LogLevel.Information);

            }
            catch (Exception ex)
            {
                LogMessage("User Error: Exception trying to update profile properties for user " + UserName + "\n" + ex.Message, LogLevel.Error);
            }
        }


        /// <summary>
        /// Used to fetch user profile property from SPO.
        /// </summary>
        /// <param name="UserName"></param>
        /// <param name="PropertyName"></param>
        /// <returns></returns>
        static string GetSingleProfileProperty(string UserName, string PropertyName)
        {
            try
            {

                var peopleManager = new PeopleManager(_clientContext);

                ClientResult<string> profileProperty = peopleManager.GetUserProfilePropertyFor(UserName, PropertyName);
                _clientContext.ExecuteQuery();

                //this is the web service way of retrieving the same thing as client API. Note: valye of propertyname is not case sensitive when using web service, but does seem to be with client API
                //UPSvc.PropertyData propertyData = userProfileService.GetUserPropertyByAccountName(UserName, PropertyName);

                if (profileProperty.Value.Length > 0)
                {
                    return profileProperty.Value;
                }
                else
                {
                    LogMessage("Cannot find a value for property " + PropertyName + " for user " + UserName, LogLevel.Information);
                    return string.Empty;
                }


            }
            catch (Exception ex)
            {
                LogMessage("User Error: Exception trying to get profile property " + PropertyName + " for user " + UserName + "\n" + ex.Message, LogLevel.Error);
                return string.Empty;
            }

        }

        /// <summary>
        /// Get multiple properties from SPO in a single call to the service
        /// </summary>
        /// <param name="UserName"></param>
        /// <param name="PropertyNames"></param>
        /// <returns></returns>
        static string[] GetMultipleProfileProperties(string UserName, string[] PropertyNames)
        {
            try
            {

                var peopleManager = new PeopleManager(_clientContext);


                UserProfilePropertiesForUser profilePropertiesForUser = new UserProfilePropertiesForUser(_clientContext, UserName, PropertyNames);
                IEnumerable<string> profilePropertyValues = peopleManager.GetUserProfilePropertiesFor(profilePropertiesForUser);

                // Load the request and run it on the server.
                _clientContext.Load(profilePropertiesForUser);
                _clientContext.ExecuteQuery();

                //convert to array and return
                return profilePropertyValues.ToArray();


            }
            catch (Exception ex)
            {
                LogMessage("Exception trying to get profile properties for user " + UserName + "\n" + ex.Message, LogLevel.Error);
                return null;
            }

        }


        /// <summary>
        /// Creates a SP Client object using SPO admin credentials, and saves client object into global variable _clientContext. Not used in the code as provided, but would be if you decide to use some of the get property information
        /// </summary>
        /// <returns></returns>
        static bool InitializeClientService()
        {

            try
            {

                LogMessage("Initializing service object for SPO Client API " + _profileSiteUrl, LogLevel.Information);
                _clientContext = new ClientContext(_profileSiteUrl);
                SecureString securePassword = GetSecurePassword(Constants._sPoAuthPasword);
                _clientContext.Credentials = new SharePointOnlineCredentials(Constants._sPoAuthUserName, securePassword);

                //LogMessage("Finished creating service object for SPO Client API " + _profileSiteUrl, LogLevel.Information);
                return true;
            }
            catch (Exception ex)
            {
                LogMessage("Error creating client context for SPO " + _profileSiteUrl + " " + ex.Message, LogLevel.Error);
                return false;
            }
        }


        /// <summary>
        /// No SPO client API for administering user profiles, so need to use traditional ASMX service for user profile work. This function initiates the 
        /// web service end point, and authenticates using Office 365 auth ticket. Use SharePointOnlineCredentials to assist with this auth.
        /// </summary>
        /// <returns></returns>
        static bool InitializeWebService()
        {
            try
            {
                string webServiceExt = "_vti_bin/userprofileservice.asmx";
                string adminWebServiceUrl = string.Empty;

                //append the web service (ASMX) url onto the admin web site URL
                if (_profileSiteUrl.EndsWith("/"))
                    adminWebServiceUrl = _profileSiteUrl + webServiceExt;
                else
                    adminWebServiceUrl = _profileSiteUrl + "/" + webServiceExt;

                LogMessage("Initializing SPO web service " + adminWebServiceUrl, LogLevel.Information);

                //get secure password from clear text password
                SecureString securePassword = GetSecurePassword(Constants._sPoAuthPasword);

                //get credentials from SP Client API, used later to extract auth cookie, so can replay to web services
                SharePointOnlineCredentials onlineCred = new SharePointOnlineCredentials(Constants._sPoAuthUserName, securePassword);


                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 |
                                       SecurityProtocolType.Tls11 |
                                       SecurityProtocolType.Tls |
                                       SecurityProtocolType.Ssl3;


                // Get the authentication cookie by passing the url of the admin web site 
                string authCookie = onlineCred.GetAuthenticationCookie(new Uri(_profileSiteUrl));

                // Create a CookieContainer to authenticate against the web service 
                CookieContainer authContainer = new CookieContainer();

                // Put the authenticationCookie string in the container 
                authContainer.SetCookies(new Uri(_profileSiteUrl), authCookie);

                // Setting up the user profile web service 
                _userProfileService = new UPSvc.UserProfileService();

                // assign the correct url to the admin profile web service 
                _userProfileService.Url = adminWebServiceUrl;

                // Assign previously created auth container to admin profile web service 
                _userProfileService.CookieContainer = authContainer;
                // LogMessage("Finished creating service object for SPO Web Service " + adminWebServiceUrl, LogLevel.Information);
                return true;
            }
            catch (Exception ex)
            {
                LogMessage("Error initiating connection to profile web service in SPO " + ex.Message, LogLevel.Error);
                return false;

            }


        }



        /// <summary>
        /// Convert clear text password into secure string
        /// </summary>
        /// <param name="Password"></param>
        /// <returns></returns>
        static SecureString GetSecurePassword(string Password)
        {
            SecureString sPassword = new SecureString();
            foreach (char c in Password.ToCharArray()) sPassword.AppendChar(c);
            return sPassword;
        }



    }
}


