using Kingdee.BOS;
using Kingdee.BOS.App;
using Kingdee.BOS.App.Data;
using Kingdee.BOS.Contracts;
using Kingdee.BOS.Core.Bill.PlugIn;
using Kingdee.BOS.Core.DynamicForm;
using Kingdee.BOS.Core.DynamicForm.Operation;
using Kingdee.BOS.Core.List;
using Kingdee.BOS.Core.Metadata;
using Kingdee.BOS.Core.Metadata.ConvertElement.ServiceArgs;
using Kingdee.BOS.Core.Metadata.FieldElement;
using Kingdee.BOS.Core.Msg;
using Kingdee.BOS.JSON;
using Kingdee.BOS.Msg;
using Kingdee.BOS.Orm;
using Kingdee.BOS.Orm.Drivers;
using Kingdee.BOS.ServiceHelper;
using Kingdee.BOS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using DynamicObject = Kingdee.BOS.Orm.DataEntity.DynamicObject;
using System.Security.Cryptography;
using Kingdee.BOS.Orm.DataEntity;
using Kingdee.BOS.Core;
using Kingdee.BOS.Core.Report;
using Kingdee.BOS.Model.ReportFilter;
using Kingdee.BOS.Core.Bill;
using Kingdee.BOS.Core.DynamicForm.PlugIn;
using Kingdee.BOS.Core.DynamicForm.PlugIn.Args;
using Kingdee.BOS.Core.Metadata.FormElement;

namespace YLZ.K3.SERVER.POST.INTERFACE
{
    [Kingdee.BOS.Util.HotUpdate]
    public  class BillTools
    {
        /// <summary>
    /// 加载单据
    /// </summary>
    /// <param name="billkey">单据标识</param>
    /// <param name="id">单据内码</param>
    /// <param name="ctx">上下文</param>
    /// <returns></returns>
        public DynamicObject  Loadbill  (string billkey,int id, Context ctx )
        {
            IMetaDataService metadataService = Kingdee.BOS.App.ServiceHelper.GetService<IMetaDataService>();
            //获取保存，加载单据服务服务
          //  ISaveService saveService = Kingdee.BOS.App.ServiceHelper.GetService<ISaveService>();
            IViewService viewService = Kingdee.BOS.App.ServiceHelper.GetService<IViewService>();
            //获取元数据
            FormMetadata materialMetadata = metadataService.Load(ctx, billkey) as FormMetadata;
            DynamicObject[] objs = viewService.Load(
                                  ctx,
                                  new object[] { id },
                                  materialMetadata.BusinessInfo.GetDynamicObjectType());
            return objs[0];


        }
        public DynamicObject Loadbill(string billkey, string id, Context ctx)
        {
            IMetaDataService metadataService = Kingdee.BOS.App.ServiceHelper.GetService<IMetaDataService>();
            //获取保存，加载单据服务服务
            //  ISaveService saveService = Kingdee.BOS.App.ServiceHelper.GetService<ISaveService>();
            IViewService viewService = Kingdee.BOS.App.ServiceHelper.GetService<IViewService>();
            //获取元数据
            FormMetadata materialMetadata = metadataService.Load(ctx, billkey) as FormMetadata;
            DynamicObject[] objs = viewService.Load(
                                  ctx,
                                  new object[] { id },
                                  materialMetadata.BusinessInfo.GetDynamicObjectType());
            return objs[0];


        }
        /// <summary>
        /// 保存单据（单据内码）
        /// </summary>
        /// <param name="billkey">单据标识</param>
        /// <param name="id">单据内码</param>
        /// <param name="ctx">上下文</param>
        /// <returns></returns>
        public IOperationResult Savebill(string billkey, int id, Context ctx)
        {
            IMetaDataService metadataService = Kingdee.BOS.App.ServiceHelper.GetService<IMetaDataService>();
            //获取保存，加载单据服务服务
           ISaveService saveService = Kingdee.BOS.App.ServiceHelper.GetService<ISaveService>();
            IViewService viewService = Kingdee.BOS.App.ServiceHelper.GetService<IViewService>();
            FormMetadata materialMetadata = metadataService.Load(ctx, billkey) as FormMetadata;
            DynamicObject[] momata = viewService.Load(
                            ctx,
                            new object[] { id },
                            materialMetadata.BusinessInfo.GetDynamicObjectType());
            var result = saveService.Save(ctx, materialMetadata.BusinessInfo, momata);
            return result;
        }
        /// <summary>
        /// 提交单据
        /// </summary>
        /// <param name="billkey">单据标识</param>
        /// <param name="id">单据内码</param>
        /// <param name="ctx">上下文</param>
        /// <returns></returns>
        public bool Submitbill(string billkey, int id, Context ctx)
        {
            IMetaDataService metadataService = Kingdee.BOS.App.ServiceHelper.GetService<IMetaDataService>();
            //获取保存，加载单据服务服务
            //ISaveService saveService = Kingdee.BOS.App.ServiceHelper.GetService<ISaveService>();
            //IViewService viewService = Kingdee.BOS.App.ServiceHelper.GetService<IViewService>();
            FormMetadata materialMetadata = metadataService.Load(ctx, billkey) as FormMetadata;
            var result = BusinessDataServiceHelper.Submit(
                       ctx,
                       materialMetadata.BusinessInfo,
                       new object[] { id },
                       "Submit");
            return result.IsSuccess;
        }
        /// <summary>
        /// 保存单据（数据包）
        /// </summary>
        /// <param name="billkey">单据标识</param>
        /// <param name="data">单据数据包</param>
        /// <param name="ctx">上下文</param>
        /// <returns></returns>
        public IOperationResult Savebilldata(string billkey, DynamicObject data, Context ctx)
        {
            IMetaDataService metadataService = Kingdee.BOS.App.ServiceHelper.GetService<IMetaDataService>();
            //获取保存，加载单据服务服务
            ISaveService saveService = Kingdee.BOS.App.ServiceHelper.GetService<ISaveService>();

            FormMetadata materialMetadata = metadataService.Load(ctx, billkey) as FormMetadata;
            DynamicObject[] datas = new DynamicObject[] { data };
            var result = saveService.Save(ctx, materialMetadata.BusinessInfo, datas);
            return result;
        }
        /// <summary>
        /// 审核单据
        /// </summary>
        /// <param name="billkey">单据标识</param>
        /// <param name="id">单据内码</param>
        /// <param name="ctx">上下文</param>
        /// <returns></returns>
        public bool Auditbill(string billkey, int id, Context ctx)
        {
            IMetaDataService metadataService = Kingdee.BOS.App.ServiceHelper.GetService<IMetaDataService>();
            //获取保存，加载单据服务服务
            //ISaveService saveService = Kingdee.BOS.App.ServiceHelper.GetService<ISaveService>();
            //IViewService viewService = Kingdee.BOS.App.ServiceHelper.GetService<IViewService>();
            FormMetadata materialMetadata = metadataService.Load(ctx, billkey) as FormMetadata;
            OperateOption auditOption = OperateOption.Create();
            var result = BusinessDataServiceHelper.Audit(ctx, materialMetadata.BusinessInfo, new object[] { id }, auditOption);
            return result.IsSuccess;
        }
        /// <summary>
        /// 获取单据数据包
        /// </summary>
        /// <param name="billkey"></param>
        /// <param name="ctx"></param>
        /// <returns></returns>
        public DynamicObject Getobj(string billkey, Context ctx)
        {
            IMetaDataService metaProxy = Kingdee.BOS.App.ServiceHelper.GetService<IMetaDataService>();
            Kingdee.BOS.Core.Metadata.FormMetadata attachmentMetadata = metaProxy.Load(ctx, billkey) as Kingdee.BOS.Core.Metadata.FormMetadata;
            BusinessInfo attachmentInfo = attachmentMetadata.BusinessInfo;
            DynamicObject attachmentData = new DynamicObject(attachmentInfo.GetDynamicObjectType());
            return attachmentData;
        }
        /// <summary>
        /// 获取单据体行数据包
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="entitykey"></param>
        /// <returns></returns>
        public DynamicObject Getentityobj(DynamicObject obj, string entitykey)
        {
            DynamicObjectCollection sktjenys = obj[entitykey] as DynamicObjectCollection;
            DynamicObject row = new DynamicObject(sktjenys.DynamicCollectionItemPropertyType);
            return row;
        }


        /// <summary>
        /// 下推操作
        /// </summary>
        /// <param name="sourceFormId">源单标识</param>
        /// <param name="targetFormId">目标单标识</param>
        /// <param name="fids">源单内码集合</param>
        /// <param name="ctx">上下文</param>
        /// <param name="rulekey">转换规则</param>
        /// <param name="fbilltype">目标单据类型</param>
        public void DoPush(string sourceFormId, string targetFormId, List<DynamicObject> fids, Context ctx, string rulekey, string fbilltype)
        {
            // 获取源单与目标单的转换规则
            IConvertService convertService = ServiceHelper.GetService<IConvertService>();
            var rules = convertService.GetConvertRules(ctx, sourceFormId, targetFormId);

            if (rules == null || rules.Count == 0)
            {
                return;
            }
            // 取勾选了默认选项的规则
            var rule = rules.FirstOrDefault(t => t.IsDefault);//取默认值，找到符合要求的赋值

            foreach (var item in rules)
            {
                if (item.Id == rulekey)
                {
                    rule = item;
                }
            }
            //var rule = rules.FirstOrDefault(t => t.IsDefault);
            //// 如果无默认规则，则取第一个
            //if (rule == null)
            //{
            //    rule = rules[2];
            //}
            List<ListSelectedRow> srcSelectedRows = new List<ListSelectedRow>();

            foreach (var fid in fids)
            {


                srcSelectedRows.Add(new ListSelectedRow(fid["Id"].ToString(), string.Empty, 0, sourceFormId));

            }
            string targetBillTypeId = fbilltype;//目标单据类型ID
            long targetOrgId = 0;
            PushArgs pushArgs = new PushArgs(rule, srcSelectedRows.ToArray())
            {
                TargetBillTypeId = targetBillTypeId,
                TargetOrgId = targetOrgId,

            };
            ConvertOperationResult operationResult = convertService.Push(ctx, pushArgs, OperateOption.Create());
            DynamicObject[] targetBillObjs = (from p in operationResult.TargetDataEntities select p.DataEntity).ToArray();
            if (targetBillObjs == null)
            {
                return;
            }
            else
            {
                //long orgId = ctx.CurrentOrganizationInfo.ID;//组织ID
                //long acctBookId = 0;
                //string parameterObjId = "AP_SystemParameter";
                //string parameterName = "F_BGJ_Month";
                //int opresult = (int)SystemParameterServiceHelper.GetParamter(ctx, orgId, acctBookId, parameterObjId, parameterName);
                //if (opresult == 1)
                //{
                //IMetaDataService metadataService = Kingdee.BOS.App.ServiceHelper.GetService<IMetaDataService>();
                ////获取保存服务
                //ISaveService saveService = Kingdee.BOS.App.ServiceHelper.GetService<ISaveService>();
                ////获取加载数据服务
                //IViewService viewService = Kingdee.BOS.App.ServiceHelper.GetService<IViewService>();
                ////获取目标单
                //FormMetadata materialMetadata = metadataService.Load(ctx, targetFormId) as FormMetadata;
                //foreach (var targetBillObj in targetBillObjs)
                //{
                //    DynamicObject[] targetobj = new DynamicObject[] { targetBillObj };
                //    var resulta = saveService.Save(ctx, materialMetadata.BusinessInfo, targetobj);//保存
                //    if (resulta.IsSuccess == false)
                //    {
                //        var draftResult = Kingdee.BOS.ServiceHelper.BusinessDataServiceHelper.Draft(ctx, materialMetadata.BusinessInfo, targetBillObj);//调用暂存
                //    }
                //}

                // }
                //if (opresult == 2)
                //{
                IMetaDataService metadataService = Kingdee.BOS.App.ServiceHelper.GetService<IMetaDataService>();
                //获取保存服务
                ISaveService saveService = Kingdee.BOS.App.ServiceHelper.GetService<ISaveService>();
                //获取加载数据服务
                IViewService viewService = Kingdee.BOS.App.ServiceHelper.GetService<IViewService>();
                //获取目标单
                FormMetadata materialMetadata = metadataService.Load(ctx, targetFormId) as FormMetadata;
                foreach (var targetBillObj in targetBillObjs)
                {
                    DynamicObject[] targetobj = new DynamicObject[] { targetBillObj };
                    var resulta = saveService.Save(ctx, materialMetadata.BusinessInfo, targetobj);//保存

                    if (resulta.IsSuccess == true)
                    {
                        var results = BusinessDataServiceHelper.Submit(
                        ctx,
                        materialMetadata.BusinessInfo,
                        new object[] { targetBillObj["Id"] },
                        "Submit");
                        if (results.IsSuccess == true)
                        {

                            OperateOption auditOption = OperateOption.Create();
                            var resultad = BusinessDataServiceHelper.Audit(ctx, materialMetadata.BusinessInfo, new object[] { targetBillObj["Id"] }, auditOption);
                        }






                    }



                    else
                    {
                        var draftResult = Kingdee.BOS.ServiceHelper.BusinessDataServiceHelper.Draft(ctx, materialMetadata.BusinessInfo, targetBillObj);//调用暂存


                    }

                }
         





            }




        }
        /// <summary>
        /// 不校验保存
        /// </summary>
        /// <param name="ctx">上下文</param>
        /// <param name="obj">单据数据包</param>
        public void Save(Context ctx,DynamicObject obj)
        {
            Kingdee.BOS.ServiceHelper.BusinessDataServiceHelper.Save(ctx, obj);
        }

        /// <summary>
        ///  给辅助资料字段赋值 
        /// </summary>
        /// <param name="service"></param>
        /// <param name="data"></param>
        /// <param name="bdfield"></param>
        /// <param name="dyValue"></param>
        /// <param name="context"></param>
        /// <param name="FzZlId"></param>
        public  void SetFZZLDataValue( DynamicObject data, BaseDataField bdfield, ref DynamicObject dyValue, Context context, string FzZlId)
        {

            IViewService service = ServiceHelper.GetService<IViewService>();
            if (FzZlId == "")
            {
                //bdfield.RefIDDynamicProperty.SetValue(data, 0);
                //bdfield.DynamicProperty.SetValue(data, null);
                dyValue = null;
            }
            else
            {
                if (dyValue == null)
                {
                    dyValue = service.LoadSingle(context, FzZlId, bdfield.RefFormDynamicObjectType);
                }
                if (dyValue != null)
                {
                    bdfield.RefIDDynamicProperty.SetValue(data, FzZlId);
                    bdfield.DynamicProperty.SetValue(data, dyValue);
                }
                else
                {
                    bdfield.RefIDDynamicProperty.SetValue(data, 0);
                    bdfield.DynamicProperty.SetValue(data, null);
                }
            }


        }
        /// <summary>
        /// 基础资料赋值（单据转换插件中）
        /// </summary>
        /// <param name="fmate">字段标识</param>
        /// <param name="fid">字段内码</param>
        /// <param name="data">当前行数据包</param>
        /// <param name="ctx">上下文</param>
        public  void SetBaseDataValue(BaseDataField fmate,long fid,DynamicObject data, Context ctx)
        {
          

            IViewService viewService = ServiceHelper.GetService<IViewService>();
            DynamicObject[] stockObjs = viewService.LoadFromCache(ctx,
            new object[] { fid },
            fmate.RefFormDynamicObjectType);
            fmate.RefIDDynamicProperty.SetValue(data, fid);
            fmate.DynamicProperty.SetValue(data, stockObjs[0]);


        }
        /// <summary>
        /// 数据库查询
        /// </summary>
        /// <param name="sql">查询sql</param>
        /// <param name="ctx">上下文</param>
        /// <returns></returns>
        public DataTable Select(string sql, Context ctx)
        {
            DataSet table = DBServiceHelper.ExecuteDataSet(ctx, sql);
            if (!(table != null && table.Tables.Count > 0 && table.Tables[0].Rows.Count > 0))
            {
                return null;
            }
            else
            {
                return table.Tables[0];
            }
        }
        /// <summary>
        /// 给用户发送消息
        /// </summary>
        /// <param name="ctx">上下文</param>
        /// <param name="formId">表单标识</param>
        /// <param name="billId">单据内码</param>
        /// <param name="title">标题</param>
        /// <param name="content">内容</param>
        /// <param name="now">时间</param>
        /// <param name="receiverId">收到消息用户内码</param>
        /// <returns></returns>
        public Message SendMessageToAimUser(Context ctx, string formId, string billId, string title, string content, DateTime now, long receiverId)
        {
            string messageId = SequentialGuid.NewGuid().ToString();
            Message msg = new DynamicObject(Message.MessageDynamicObjectType);
            msg.MsgType = MsgType.WorkflowMessage;//流程消息  
            msg.MessageId = messageId;
            msg.Title = title;
            msg.Content = content;
            msg.CreateTime = now;
            msg.SenderId = ctx.UserId;//
            msg.ObjectTypeId = formId;
            msg.KeyValue = billId;
            msg.ReceiverId = receiverId;
            IDbDriver driver = new OLEDbDriver(ctx);
            var dataManager = Kingdee.BOS.Orm.DataManagerUtils.GetDataManager(Message.MessageDynamicObjectType, driver);
            dataManager.Save(msg.DataEntity);

            return msg;
        }
        /// <summary>
        /// 连接sqlserver数据库执行查询语句
        /// </summary>
        /// <param name="serverip">数据库地址</param>
        /// <param name="database">数据库名称</param>
        /// <param name="userid">登录账户</param>
        /// <param name="password">登录密码</param>
        /// <param name="sql">查询sql语句</param>
        /// <returns></returns>
        public DataSet Connectsqlserver(string serverip,string database,string userid,string password,string sql)
        {
            SqlConnection opensqlserver = new SqlConnection(string.Format("Server={0};Database={1};uid={2};pwd={3}",serverip,database,userid,password,sql));
            opensqlserver.Open();
            SqlCommand execsql = new SqlCommand();
            execsql.CommandText = sql;
            execsql.Connection = opensqlserver;
            SqlDataAdapter dbAdapter = new SqlDataAdapter(execsql);
            DataSet table = new DataSet();
            dbAdapter.Fill(table);
            opensqlserver.Close();
            return table;
        }
        /// <summary>
        /// webclient下载文件
        /// </summary>
        /// <param name="url">文件路径</param>
        /// <param name="download">下载到哪个路径</param>
        /// <returns></returns>
        public  string DownloadFile(string url,string download)
        {
            try
            {
                string fileName = CreateFileName(url);
                if (!Directory.Exists(download))
                {
                    Directory.CreateDirectory(download);
                }
                WebClient client = new WebClient();
                client.DownloadFile(url, download + fileName);
                return fileName;
            }
            catch
            {
                return "下载失败";
            }
        }
        /// <summary>
        /// webclient下载创建文件名称
        /// </summary>
        /// <param name="url">文件路径</param>
        /// <returns></returns>
        public  string CreateFileName(string url)
        {
            string fileName = "";
            string fileExt = url.Substring(url.LastIndexOf(@"\"));
            fileName = fileExt;
            return fileName;
        }
        /// <summary>
        /// 金蝶云上传附件
        /// </summary>
        /// <param name="fileByte">文件大小</param>
        /// <param name="file">文件</param>
        /// <param name="ctx">上下文</param>
        /// <returns></returns>
        public string Upload(byte[] fileByte, FileInfo file, Context ctx,string loginurl)
        {
            bool bResult = false;
            string sUrl = string.Format("{0}FileUpLoadServices/FileService.svc/Upload2Attachment/?fileName={1}&fileid={2}&token={5}&dbid={3}&last={4}",
                   loginurl+"/", file.Name, string.Empty, ctx.DBId, true, ctx.UserToken);
            using (WebClient webClient = new WebClient() { Encoding = Encoding.UTF8 })
            {
                webClient.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1)");
                webClient.Headers.Add("Charsert", "UTF-8");
                webClient.Headers.Add("Content-Type", "multipart/form-data;");
                byte[] responseBytes;

                try
                {
                    responseBytes = webClient.UploadData(sUrl, "POST", fileByte);
                    JSONObject jObj = JSONObject.Parse(Encoding.UTF8.GetString(responseBytes));
                    JSONObject uploadResult = jObj["Upload2AttachmentResult"] as JSONObject;
                    bResult = Convert.ToBoolean(uploadResult["Success"].ToString());
                    if (bResult)
                    {
                        return uploadResult["FileId"].ToString();
                    }
                    else
                    {

                        return string.Empty;
                    }
                }
                catch (WebException ex)
                {
                    Stream resp = ex.Response.GetResponseStream();
                    responseBytes = new byte[ex.Response.ContentLength];
                    resp.Read(responseBytes, 0, responseBytes.Length);
                    resp.Close();
                    resp.Dispose();

                    return "上传失败";
                }



            }
        }
        /// <summary>
        /// http下载文件
        /// </summary>
        /// <param name="url">下载文件地址</param>
        /// <param name="path">文件存放地址，包含文件名</param>
        /// <returns></returns>
        public string HttpDownload(string url, string path)
        {
            string tempPath = System.IO.Path.GetDirectoryName(path) + @"\temp";
            System.IO.Directory.CreateDirectory(tempPath);  //创建临时文件目录
            string tempFile = tempPath + @"\" + System.IO.Path.GetFileName(path) + ".temp"; //临时文件
            if (System.IO.File.Exists(tempFile))
            {
                System.IO.File.Delete(tempFile);    //存在则删除
            }
            try
            {
                FileStream fs = new FileStream(tempFile, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
                // 设置参数
                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                //发送请求并获取相应回应数据
                HttpWebResponse response = request.GetResponse() as HttpWebResponse;
                //直到request.GetResponse()程序才开始向目标网页发送Post请求
                Stream responseStream = response.GetResponseStream();
                //创建本地文件写入流
                //Stream stream = new FileStream(tempFile, FileMode.Create);
                byte[] bArr = new byte[1024];
                int size = responseStream.Read(bArr, 0, (int)bArr.Length);
                while (size > 0)
                {
                    //stream.Write(bArr, 0, size);
                    fs.Write(bArr, 0, size);
                    size = responseStream.Read(bArr, 0, (int)bArr.Length);
                }
                //stream.Close();
                fs.Close();
                responseStream.Close();
                System.IO.File.Move(tempFile, path);
                return "下载成功";
            }
            catch (Exception ex)
            {
                
                return ex.Message;
            }
        }
        /// <summary>
        /// http下载文件创建文件名称
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public string HttpCreateFileName(string url)
        {
            string fileName = "";
            string fileExt = url.Substring(url.LastIndexOf(@"/")).Replace("/","");
            fileName = fileExt;
            return fileName;
        }
      

        /// <summary>
        /// POST方式调用接口 data 参数，url接口地址，返回json字符串
        /// </summary>
        /// <param name="data"></param>
        /// <param name="url"></param>
        /// <returns></returns>
        public string Post(string data, string url)
        {
            //先根据用户请求的uri构造请求地址
            string serviceUrl = string.Format("{0}", url);
            //创建Web访问对象
            HttpWebRequest myRequest = (HttpWebRequest)WebRequest.Create(serviceUrl);
            //把用户传过来的数据转成“UTF-8”的字节流
            byte[] buf = System.Text.Encoding.GetEncoding("UTF-8").GetBytes(data);

            myRequest.Method = "POST";
            myRequest.ContentLength = buf.Length;
            myRequest.ContentType = "application/json";
            myRequest.MaximumAutomaticRedirections = 1;
            myRequest.AllowAutoRedirect = true;
            //发送请求
            Stream stream = myRequest.GetRequestStream();
            stream.Write(buf, 0, buf.Length);
            stream.Close();

            //获取接口返回值
            //通过Web访问对象获取响应内容
            HttpWebResponse myResponse = (HttpWebResponse)myRequest.GetResponse();
            //通过响应内容流创建StreamReader对象，因为StreamReader更高级更快
            StreamReader reader = new StreamReader(myResponse.GetResponseStream(), Encoding.UTF8);
            //string returnXml = HttpUtility.UrlDecode(reader.ReadToEnd());//如果有编码问题就用这个方法
            string returnXml = reader.ReadToEnd();//利用StreamReader就可以从响应内容从头读到尾
            reader.Close();
            myResponse.Close();
            return returnXml;

        }
     /// <summary>
     /// ASE加密 加密字符串，密匙，类型EBC,填充PKCS7
     /// </summary>
     /// <param name="EncryptStr"></param>
     /// <param name="Key"></param>
     /// <returns></returns>
        public string AesEncrypt(string EncryptStr, string Key)
        {
            try
            {
                byte[] keyArray = Encoding.UTF8.GetBytes(Key);
                //  byte[] keyArray = Convert.FromBase64String(Key);
                byte[] toEncryptArray = Encoding.UTF8.GetBytes(EncryptStr);

                RijndaelManaged rDel = new RijndaelManaged();
                rDel.Key = keyArray;
                rDel.Mode = CipherMode.ECB;
                rDel.Padding = PaddingMode.PKCS7;
                rDel.KeySize = 128;
                rDel.BlockSize = 128;
                ICryptoTransform cTransform = rDel.CreateEncryptor();
                byte[] resultArray = cTransform.TransformFinalBlock(toEncryptArray, 0, toEncryptArray.Length);

                return Convert.ToBase64String(resultArray, 0, resultArray.Length);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
     
     /// <summary>
     /// ASE解密ECB模式
     /// </summary>
     /// <param name="DecryptStr"></param>
     /// <param name="Key"></param>
     /// <returns></returns>
        public string AESDecrypt(string DecryptStr, string Key)
        {
            try
            {
                byte[] keyArray = Encoding.UTF8.GetBytes(Key);
                // byte[] keyArray = Convert.FromBase64String(Key);
                byte[] toEncryptArray = Convert.FromBase64String(DecryptStr);

                RijndaelManaged rDel = new RijndaelManaged();
                rDel.Key = keyArray;
                rDel.Mode = CipherMode.ECB;
                rDel.Padding = PaddingMode.Zeros;
                rDel.BlockSize = 128;
              //  rDel.KeySize = 128;
                ICryptoTransform cTransform = rDel.CreateDecryptor();
                byte[] resultArray = cTransform.TransformFinalBlock(toEncryptArray, 0, toEncryptArray.Length);

                return Encoding.UTF8.GetString(resultArray);//  UTF8Encoding.UTF8.GetString(resultArray);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        /// <summary>
        /// ASE加密 CBC模式
        /// </summary>
        /// <param name="text"></param>
        /// <param name="password"></param>
        /// <param name="iv"></param>
        /// <returns></returns>
        public  string PYAESEncrypt(string text, string password, string iv)
        {
            RijndaelManaged rijndaelCipher = new RijndaelManaged();
            rijndaelCipher.Mode = CipherMode.CBC;
            rijndaelCipher.Padding = PaddingMode.Zeros;
         //   rijndaelCipher.KeySize = 128;
            rijndaelCipher.BlockSize = 128;
            byte[] pwdBytes = System.Text.Encoding.UTF8.GetBytes(password);
            byte[] keyBytes = new byte[16];
            int len = pwdBytes.Length;
            if (len > keyBytes.Length) len = keyBytes.Length;
            System.Array.Copy(pwdBytes, keyBytes, len);
            rijndaelCipher.Key = keyBytes;
            byte[] ivBytes = System.Text.Encoding.UTF8.GetBytes(iv);
            rijndaelCipher.IV = ivBytes;
            ICryptoTransform transform = rijndaelCipher.CreateEncryptor();
            byte[] plainText = Encoding.UTF8.GetBytes(text);
            byte[] cipherBytes = transform.TransformFinalBlock(plainText, 0, plainText.Length);
            return Convert.ToBase64String(cipherBytes);
        }
        /// <summary>
        /// ASE加密 调用PMS接口 数据包加密使用
        /// </summary>
        /// <param name="text"></param>
        /// <param name="password"></param>
        /// <param name="iv"></param>
        /// <returns></returns>
        public string PMSAESEncrypt(string text, string password, string iv)
        {
            RijndaelManaged rijndaelCipher = new RijndaelManaged();
            rijndaelCipher.Mode = CipherMode.CBC;
            rijndaelCipher.Padding = PaddingMode.PKCS7;
            //   rijndaelCipher.KeySize = 128;
            rijndaelCipher.BlockSize = 128;
            byte[] pwdBytes = System.Text.Encoding.UTF8.GetBytes(password);
            byte[] keyBytes = new byte[16];
            int len = pwdBytes.Length;
            if (len > keyBytes.Length) len = keyBytes.Length;
            System.Array.Copy(pwdBytes, keyBytes, len);
            rijndaelCipher.Key = keyBytes;
            byte[] ivBytes = System.Text.Encoding.UTF8.GetBytes(iv);
            rijndaelCipher.IV = ivBytes;
            ICryptoTransform transform = rijndaelCipher.CreateEncryptor();
            byte[] plainText = Encoding.UTF8.GetBytes(text);
            byte[] cipherBytes = transform.TransformFinalBlock(plainText, 0, plainText.Length);
            return Convert.ToBase64String(cipherBytes);
        }
        /// <summary>
        /// ASE解密CBC模式
        /// </summary>
        /// <param name="text"></param>
        /// <param name="password"></param>
        /// <param name="iv"></param>
        /// <returns></returns>
        public string PYAESDecrypt(string text, string password, string iv)
        {
            try
            {

                RijndaelManaged rijndaelCipher = new RijndaelManaged();
                rijndaelCipher.Mode = CipherMode.CBC;
                rijndaelCipher.Padding = PaddingMode.Zeros;
                //  rijndaelCipher.KeySize = 128;
                rijndaelCipher.BlockSize = 128;
                byte[] encryptedData = Convert.FromBase64String(text);
                byte[] pwdBytes = System.Text.Encoding.UTF8.GetBytes(password);
                byte[] keyBytes = new byte[16];
                int len = pwdBytes.Length;
                if (len > keyBytes.Length) len = keyBytes.Length;
                System.Array.Copy(pwdBytes, keyBytes, len);
                rijndaelCipher.Key = keyBytes;
                byte[] ivBytes = System.Text.Encoding.UTF8.GetBytes(iv);
                rijndaelCipher.IV = ivBytes;
                ICryptoTransform transform = rijndaelCipher.CreateDecryptor();
                byte[] plainText = transform.TransformFinalBlock(encryptedData, 0, encryptedData.Length);
                return Encoding.UTF8.GetString(plainText);
            }
            catch (Exception ex)
            {

                return ex.Message;
            }
        }
        /// <summary>
        /// 批量删除单据
        /// </summary>
        /// <param name="billkey"></param>
        /// <param name="id"></param>
        /// <param name="ctx"></param>
        /// <returns></returns>
        public bool Deletebill(string billkey, object[] id, Context ctx)
        {

            IDeleteService deleteService =
             Kingdee.BOS.App.ServiceHelper.GetService<IDeleteService>();
            //获取元数据服务
            IMetaDataService metaDataService =
                Kingdee.BOS.App.ServiceHelper.GetService<IMetaDataService>();
            FormMetadata moin =
                metaDataService.Load(ctx, billkey) as FormMetadata;
            //执行删除服务的完整过程
            var result = deleteService.Delete(
                   ctx,
                   moin.BusinessInfo,
                   id);
            return result.IsSuccess;
        }
        /// <summary>
        /// 删除单个单据
        /// </summary>
        /// <param name="billkey"></param>
        /// <param name="id"></param>
        /// <param name="ctx"></param>
        /// <returns></returns>
        public bool Deletebill(string billkey, int id, Context ctx)
        {

            IDeleteService deleteService =
             Kingdee.BOS.App.ServiceHelper.GetService<IDeleteService>();
            //获取元数据服务
            IMetaDataService metaDataService =
                Kingdee.BOS.App.ServiceHelper.GetService<IMetaDataService>();
            FormMetadata moin =
                metaDataService.Load(ctx, billkey) as FormMetadata;
            //执行删除服务的完整过程
            var result = deleteService.Delete(
                   ctx,
                   moin.BusinessInfo,
                   new object[] { id});
            return result.IsSuccess;
        }
        /// <summary>
        /// 获取到期日，ywdate业务日期，sktjeny收款条件分录数据包
        /// </summary>
        /// <param name="ywdate"></param>
        /// <param name="sktjeny"></param>
        /// <returns></returns>
        public DateTime GetDQR(DateTime ywdate, DynamicObject sktjeny)//根据收款条件计算到期日,应收单日期,收款计划单据体数据包
        {

            if (sktjeny != null)
            {

                DateTime xtdqr = ywdate;


                bool yd = Convert.ToBoolean(sktjeny["ISMONTHEND"]);//是否勾选月底
                dynamic dqrfs = sktjeny["DuecalMethodID"] as dynamic;//到期日结算方式
                if (dqrfs != null)
                {
                    if (dqrfs["Name"].ToString() == "月结")
                    {
                        int jsr = Convert.ToInt32(sktjeny["THISMONTHEND"]);//结算日
                        int ys = Convert.ToInt32(sktjeny["ODMonth"]);//月数
                        int ts = Convert.ToInt32(sktjeny["ODDay"]);//天数
                        int ydrs = Convert.ToInt32(sktjeny["ODDayToEndMonth"]);//距离月底天数
                        int rq = ywdate.Day;//收款条件日期

                        if (rq > jsr)
                        {
                            if (yd == true)
                            {
                                xtdqr = ywdate.AddMonths(ys + 3).AddDays(-ywdate.Day - ydrs);

                            }
                            else
                            {
                                xtdqr = ywdate.AddMonths(ys + 2).AddDays(-ywdate.Day + ts);

                            }

                        }
                        else
                        {
                            if (yd == true)
                            {
                                xtdqr = ywdate.AddMonths(ys + 2).AddDays(-ywdate.Day - ydrs);

                            }
                            else
                            {
                                xtdqr = ywdate.AddMonths(ys + 1).AddDays(-ywdate.Day + ts);

                            }

                        }

                    }

                    if (dqrfs["Name"].ToString() == "X天后")
                    {
                        int ys = Convert.ToInt32(sktjeny["ODMonth"]);//月数
                        int ts = Convert.ToInt32(sktjeny["ODDay"]);//天数
                        int ydrs = Convert.ToInt32(sktjeny["ODDayToEndMonth"]);//距离月底天数

                        if (yd == true)
                        {
                            xtdqr = ywdate.AddMonths(ys + 2).AddDays(-ywdate.Day - ydrs);

                        }
                        else
                        {
                            xtdqr = ywdate.AddMonths(ys).AddDays(ts);

                        }
                    }

                    if (dqrfs["Name"].ToString() == "固定日")
                    {

                        int ys = Convert.ToInt32(sktjeny["THISMONTHEND"]);//月数
                        int ts = Convert.ToInt32(sktjeny["ODDay"]);//天数
                        int gd1 = Convert.ToInt32(sktjeny["ODFixedDayOne"]);//固定1
                        int gd2 = Convert.ToInt32(sktjeny["ODFixedDayTwo"]);//固定2
                        int gd3 = Convert.ToInt32(sktjeny["ODFixedDayThree"]);//固定3
                        int rq = ywdate.Day;//收款条件日期


                        if (rq + ts <= gd1)
                        {
                            xtdqr = ywdate.AddMonths(ys + 1).AddDays(-ywdate.Day + gd1);

                        }
                        if (gd1 < rq + ts && rq + ts <= gd2)
                        {
                            xtdqr = ywdate.AddMonths(ys + 1).AddDays(-ywdate.Day + gd2);

                        }
                        else
                        {
                            xtdqr = ywdate.AddMonths(ys + 1).AddDays(-ywdate.Day + gd3);

                        }

                    }
                    if (dqrfs["Name"].ToString() == "交易日")
                    {
                        xtdqr = ywdate;
                    }


                }
                else
                {
                    xtdqr = ywdate;
                }


                return xtdqr;


            }
            else
            {
                return ywdate;
            }


        }
        /// <summary>
        /// 调用执行计划ctx 上下文,runid执行计划内码
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="runid"></param>
        public void Dyrun(Context ctx, string runid)

        {
            Schedule scheduleByTypeId = ScheduleBusinessServiceHelper.GetScheduleByTypeId(ctx, runid);
                    if (scheduleByTypeId != null && !string.IsNullOrEmpty(scheduleByTypeId.ScheduleId))
                    {
                        
                        scheduleByTypeId.IsDebug = true;
                      //  scheduleByTypeId.Parameters = "1";
                        PaaSScheduleServiceHelper.RunSchedule(ctx, scheduleByTypeId);
                     
                    }
}

       public string CallHttp( string  url, string method, string param)
        {


            HttpWebRequest http = (HttpWebRequest)WebRequest.Create(url);
            http.Method = method;
            http.ContentType = "application/json";
            Stream req = http.GetRequestStream();
            byte[]   data = UnicodeEncoding.UTF8.GetBytes(param);
            req.Write(data, 0, data.Length);
            req.Flush();
            HttpWebResponse myResponse = (HttpWebResponse)http.GetResponse();
            //通过响应内容流创建StreamReader对象，因为StreamReader更高级更快
            StreamReader reader = new StreamReader(myResponse.GetResponseStream(), Encoding.UTF8);
            //string returnXml = HttpUtility.UrlDecode(reader.ReadToEnd());//如果有编码问题就用这个方法
            string returnXml = reader.ReadToEnd();//利用StreamReader就可以从响应内容从头读到尾
            reader.Close();
            myResponse.Close();
            return returnXml;

        }

        /// <summary>
        /// 此方法模拟人为查询物料收发明细表数据返回DataTable拿到DataTable自行操作DataTable获取需要数据，ctx上下文，zzid 组织编码，starttime开始时间，endtime结束时间，fma物料编码，fstock仓库编码，flot批号，billstate单据状态
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="zzid"></param>
        /// <param name="yf"></param>
        /// <param name="year"></param>
        /// <param name="billstate"></param>
        /// <returns></returns>
        public DataTable Wlsfhzb(Context ctx, string zzid, string starttime, string endtime, int fmaid, int fstockid, string flot, string billstate,DynamicObject fmadata,DynamicObject fstockdata)
        {
            DataTable temptable = null;
           
                ISysReportService sysReporSservice = ServiceFactory.GetSysReportService(ctx);
                IPermissionService permissionService = ServiceFactory.GetPermissionService(ctx);
                var filterMetadata = FormMetaDataCache.GetCachedFilterMetaData(ctx);//加载字段比较条件元数据。
                var reportMetadata = FormMetaDataCache.GetCachedFormMetaData(ctx, "STK_StockDetailRpt");//物料收发汇总表标识
                var reportFilterMetadata = FormMetaDataCache.GetCachedFormMetaData(ctx, "STK_StockDetailFilter");//物料收发汇总表过滤框
                var reportFilterServiceProvider = reportFilterMetadata.BusinessInfo.GetForm().GetFormServiceProvider();
                var model = new SysReportFilterModel();
                model.SetContext(ctx, reportFilterMetadata.BusinessInfo, reportFilterServiceProvider);
                model.FormId = reportFilterMetadata.BusinessInfo.GetForm().Id;
                model.FilterObject.FilterMetaData = filterMetadata;
                model.InitFieldList(reportMetadata, reportFilterMetadata);
                model.GetSchemeList();

                var openParameter = new Dictionary<string, object>();
                var parameterDataFormId = reportMetadata.BusinessInfo.GetForm().ParameterObjectId;
                var parameterDataMetadata = FormMetaDataCache.GetCachedFormMetaData(ctx, parameterDataFormId);
                var parameterData = UserParamterServiceHelper.Load(ctx, parameterDataMetadata.BusinessInfo, ctx.UserId, "STK_StockDetailRpt", KeyConst.USERPARAMETER_KEY);

                var filter = model.GetFilterParameter();
                DynamicObject flter = filter.CustomFilter;
                flter["StockOrgId"] = zzid;//组织
                flter["BeginDate"] = starttime;//开始时间
                flter["EndDate"] = endtime;//结束时间
                flter["BeginMaterialId_Id"] = fmaid;
                flter["BeginMaterialId"] = fmadata;
                flter["BeginStockId_Id"] = fstockid;
                flter["BeginStockId"] = fstockdata;
                //flter["BeginLotNo"] = eny[flotorm];
                flter["BeginLotNo_Text"] = flot;
                //     flter["BeginLotNo_Id"] = eny["Lot_Id"];
                flter["EndMaterialId_Id"] = fmaid;
                flter["EndMaterialId"] = fmadata;
                flter["EndStockId_Id"] = fstockid;
                flter["EndStockId"] = fstockdata;
                //flter["EndLotNo"] = eny[flotorm];
                flter["EndLotNo_Text"] = flot;
                //   flter["EndLotNo_Id"] = eny["Lot_Id"];
                flter["BillStatus"] = billstate;//单据状态
                flter["QCPriceSource"] = "3";
                flter["IOPriceSource"] = "1";
                //   flter["TransDirectPriceSource"] = "2";//7.3版本没有改字段
                //    flter["VMIInstockPriceSource"] = "2";
                //   flter["FIsHsStockField"] = false;
                flter["IncludReceive"] = false;
                flter["ShowForbidMaterial"] = false;
                flter["NoHappenNoShow"] = true;
                flter["SplitPageByOwner"] = true;
                flter["NoQtyNoShow"] = false;
                flter["NoBalanceNoShow"] = false;
                flter["NoStockAdjust"] = false;
                flter["NoHappenNoQcNoShow"] = true;
                flter["NoInnerTrans"] = false;
                //  flter["StockQtyRule"] = "0";
                IRptParams p = new RptParams();
                p.FormId = reportFilterMetadata.BusinessInfo.GetForm().Id;
                p.StartRow = 1;
                p.EndRow = int.MaxValue;//StartRow和EndRow是报表数据分页的起始行数和截至行数，一般取所有数据，所以EndRow取int最大值。
                p.FilterParameter = filter;
                p.FilterFieldInfo = model.FilterFieldInfo;
                p.BaseDataTempTable.AddRange(permissionService.GetBaseDataTempTable(ctx, reportMetadata.BusinessInfo.GetForm().Id));
                p.ParameterData = parameterData;
                p.CustomParams.Add(KeyConst.OPENPARAMETER_KEY, openParameter);
                using (DataTable dt = sysReporSservice.GetData(ctx, reportMetadata.BusinessInfo, p))//根据拼接好的过滤条件调用系统标准报表查询出结果集
                {

                    temptable = dt;

                }
                ServiceFactory.CloseService(sysReporSservice);
                ServiceFactory.CloseService(permissionService);

                return temptable;
            
           
           
           
           


        }
        public IBillView CreateBillView(Context ctx, string key, string id)
        {
            // 读取物料的元数据
            FormMetadata meta = MetaDataServiceHelper.Load(ctx, key) as FormMetadata;
            Form form = meta.BusinessInfo.GetForm();
            // 创建用于引入数据的单据view
            Type type = Type.GetType("Kingdee.BOS.Web.Import.ImportBillView,Kingdee.BOS.Web");
            // Type type = Type.GetType("Kingdee.BOS.Web.Bill.BillView,Kingdee.BOS.Web");
            var billView = (IDynamicFormViewService)Activator.CreateInstance(type);
            // 开始初始化billView：
            // 创建视图加载参数对象，指定各种参数，如FormId, 视图(LayoutId)等
            BillOpenParameter openParam = CreateOpenParameter(ctx, meta, id);
            // 动态领域模型服务提供类，通过此类，构建MVC实例
            var provider = form.GetFormServiceProvider();
            billView.Initialize(openParam, provider);
            billView.LoadData();
            (billView as IBillView).CommitNetworkCtrl();
            return billView as IBillView;
        }
        public BillOpenParameter CreateOpenParameter(Context ctx, FormMetadata meta, string id)
        {
            Form form = meta.BusinessInfo.GetForm();
            // 指定FormId, LayoutId
            BillOpenParameter openParam = new BillOpenParameter(form.Id, meta.GetLayoutInfo().Id);
            // 数据库上下文
            openParam.Context = ctx;
            // 本单据模型使用的MVC框架
            openParam.ServiceName = form.FormServiceName;
            // 随机产生一个不重复的PageId，作为视图的标识
            openParam.PageId = Guid.NewGuid().ToString();
            // 元数据
            openParam.FormMetaData = meta;
            // 界面状态：新增 (修改、查看)


            openParam.Status = OperationStatus.EDIT;


            // 单据主键
            openParam.PkValue = id;
            // 界面创建目的：普通无特殊目的 （为工作流、为下推、为复制等）
            openParam.CreateFrom = CreateFrom.Default;
            // 基础资料分组维度：基础资料允许添加多个分组字段，每个分组字段会有一个分组维度
            // 具体分组维度Id，请参阅 form.FormGroups 属性
            openParam.GroupId = "";
            // 基础资料分组：如果需要为新建的基础资料指定所在分组，请设置此属性
            openParam.ParentId = 0;
            // 单据类型
            openParam.DefaultBillTypeId = "";
            // 业务流程
            openParam.DefaultBusinessFlowId = "";
            // 主业务组织改变时，不用弹出提示界面
            openParam.SetCustomParameter("ShowConfirmDialogWhenChangeOrg", false);
            // 插件
            List<AbstractDynamicFormPlugIn> plugs = form.CreateFormPlugIns();
            openParam.SetCustomParameter(FormConst.PlugIns, plugs);
            PreOpenFormEventArgs args = new PreOpenFormEventArgs(ctx, openParam);
            foreach (var plug in plugs)
            {// 触发插件PreOpenForm事件，供插件确认是否允许打开界面
                plug.PreOpenForm(args);
            }
            if (args.Cancel == true)
            {// 插件不允许打开界面
             // 本案例不理会插件的诉求，继续....
            }
            // 返回

            return openParam;
        }

    }
}
