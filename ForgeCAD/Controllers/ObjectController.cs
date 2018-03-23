using Autodesk.Forge;
using Autodesk.Forge.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using ExcelLibrary.SpreadSheet;
using System.Linq;
using System.Text.RegularExpressions;


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


        [HttpPost]
        [Route("api/forge/object/excel")]
        public async Task ExcelObject([FromBody] ObjectModel objModel)
        {

            // authenticate with Forge Must have data read scope
            dynamic oauth = await OAuthController.GetInternalAsync();

            // get the user selected object
            ObjectsApi objects = new ObjectsApi();
            objects.Configuration.AccessToken = oauth.access_token;
            dynamic selectedObject = await objects.GetObjectDetailsAsync(objModel.bucketKey, objModel.objectName);

            string objectId = selectedObject.objectId;
            string objectKey = selectedObject.objectKey;
            

            string xlsFileName = objectKey.Replace(".rvt", ".xls");
            var xlsPath = Path.Combine(HttpContext.Current.Server.MapPath("~/App_Data"), objModel.bucketKey, xlsFileName);//.guid, xlsFileName);
            ///////////if (File.Exists(xlsPath))
            ///////////    return SendFile(xlsPath);// if the Excel file was already generated

            DerivativesApi derivative = new DerivativesApi();
            derivative.Configuration.AccessToken = oauth.access_token;

            // get the derivative metadata
            dynamic metadata = await derivative.GetMetadataAsync(objectId.Base64Encode());
            foreach (KeyValuePair<string, dynamic> metadataItem in new DynamicDictionaryItems(metadata.data.metadata))
            {
                dynamic hierarchy = await derivative.GetModelviewMetadataAsync(objectId.Base64Encode(), metadataItem.Value.guid);
                dynamic properties = await derivative.GetModelviewPropertiesAsync(objectId.Base64Encode(), metadataItem.Value.guid);

                Workbook xls = new Workbook();
                foreach (KeyValuePair<string, dynamic> categoryOfElements in new DynamicDictionaryItems(hierarchy.data.objects[0].objects))
                {
                    string name = categoryOfElements.Value.name;
                    Worksheet sheet = new Worksheet(name);
                    for (int i = 0; i < 100; i++) sheet.Cells[i, 0] = new Cell(""); // unless we have at least 100 cells filled, Excel understand this file as corrupted

                    List<long> ids = GetAllElements(categoryOfElements.Value.objects);
                    int row = 1;
                    foreach (long id in ids)
                    {
                        Dictionary<string, object> props = GetProperties(id, properties);
                        int collumn = 0;
                        foreach (KeyValuePair<string, object> prop in props)
                        {
                            sheet.Cells[0, collumn] = new Cell(prop.Key.ToString());
                            sheet.Cells[row, collumn] = new Cell(prop.Value.ToString());
                            collumn++;
                        }

                        row++;
                    }

                    xls.Worksheets.Add(sheet);
                }


                //Where to save the excel file
                string pathUser = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                string pathDownload = Path.Combine(pathUser, "Downloads", xlsFileName);

                //try catch save the file to the relevant place
                try
                {
                    var fstream = new System.IO.FileStream(pathDownload, FileMode.CreateNew);
                    xls.SaveToStream(fstream);
                }
                catch (Exception e)
                {
                    Debug.Print("Exception when calling ObjectsApi.DownloadObject: " + e.Message);
                }
            }
            //No need to return anything
            //return SendFile(xlsPath);
        }

        /// <summary>
        /// Model for DeleteObject method
        /// </summary>
        public class ObjectModel
        {
            public string bucketKey { get; set; }
            public string objectName { get; set; }
            public string objectKey { get; set; }
        }

        #region excelmisc
        /// <summary>
        /// Recursively run through the list of objects hierarchy getting alls IDs with no children
        /// </summary>
        /// <param name="objects"></param>
        /// <returns></returns>
        private List<long> GetAllElements(dynamic objects)
        {
            List<long> ids = new List<long>();
            foreach (KeyValuePair<string, dynamic> item in new DynamicDictionaryItems(objects))
            {
                foreach (KeyValuePair<string, dynamic> keys in item.Value.Dictionary)
                {
                    if (keys.Key.Equals("objects"))
                    {
                        return GetAllElements(item.Value.objects);
                    }
                }
                foreach (KeyValuePair<string, dynamic> element in objects.Dictionary)
                {
                    if (!ids.Contains(element.Value.objectid))
                        ids.Add(element.Value.objectid);
                }

            }
            return ids;
        }

        /// <summary>
        /// Get a list of properties for a given ID
        /// </summary>
        /// <param name="id"></param>
        /// <param name="properties"></param>
        /// <returns></returns>
        private Dictionary<string, object> GetProperties(long id, dynamic properties)
        {
            Dictionary<string, object> returnProps = new Dictionary<string, object>();
            foreach (KeyValuePair<string, dynamic> objectProps in new DynamicDictionaryItems(properties.data.collection))
            {
                if (objectProps.Value.objectid != id) continue;
                string name = objectProps.Value.name;
                long elementId = long.Parse(Regex.Match(name, @"\d+").Value);
                returnProps.Add("ID", elementId);
                returnProps.Add("Name", name.Replace("[" + elementId.ToString() + "]", string.Empty));
                foreach (KeyValuePair<string, dynamic> objectPropsGroup in new DynamicDictionaryItems(objectProps.Value.properties))
                {
                    if (objectPropsGroup.Key.StartsWith("__")) continue;
                    foreach (KeyValuePair<string, dynamic> objectProp in new DynamicDictionaryItems(objectPropsGroup.Value))
                    {
                        if (!returnProps.ContainsKey(objectProp.Key))
                            returnProps.Add(objectProp.Key, objectProp.Value);
                        else
                            Debug.Write(objectProp.Key);
                    }
                }
            }
            return returnProps;
        }
        #endregion
    }
}

