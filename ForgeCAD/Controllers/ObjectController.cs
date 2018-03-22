using Autodesk.Forge;
using Autodesk.Forge.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Web.Http;

namespace ForgeCAD.Controllers
{
    public class ObjectController : ApiController
    {
        [HttpPost]
        [Route("api/forge/object/delete")]
        public async Task DeleteObject([FromBody]ObjectModel objModel)
        {
            dynamic oauth = await OAuthController.GetInternalAsync();


            var apiInstance = new ObjectsApi();
            var bucketKey = objModel.bucketKey;  // string | URL-encoded bucket key
            var objectName = objModel.objectName;  // string | URL-encoded object name

            try
            {
                apiInstance.DeleteObject(bucketKey, objectName);
            }
            catch (Exception e)
            {
                Debug.Print("Exception when calling ObjectsApi.DeleteObject: " + e.Message);
            }
        }

        [HttpPost]
        [Route("api/forge/object/download")]
        public async Task DownloadObject([FromBody]ObjectModel objModel)
        {
            dynamic oauth = await OAuthController.GetInternalAsync();


            var apiInstance = new ObjectsApi();
            string pathUser = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string pathDownload = Path.Combine(pathUser, "Downloads");
            apiInstance.Configuration.AccessToken = oauth.access_token;
            var bucketKey = objModel.bucketKey;  // string | URL-encoded bucket key
            var objectName = objModel.objectName;  // string | URL-encoded object name

            try
            {
                System.IO.Stream result = apiInstance.GetObject(bucketKey, objectName);
                var fstream = new System.IO.FileStream(Path.Combine(pathDownload, objectName), FileMode.CreateNew);
                result.CopyTo(fstream);
            }
            catch (Exception e)
            {
                Debug.Print("Exception when calling ObjectsApi.DownloadObject: " + e.Message);
            }
        }

        /// <summary>
        /// Model for DeleteObject method
        /// </summary>
        public class ObjectModel
        {
            public string bucketKey { get; set; }
            public string objectName { get; set; }
        }
    }
}

