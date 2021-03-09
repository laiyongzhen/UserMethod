using Kingdee.BOS;
using Kingdee.BOS.App;
using Kingdee.BOS.App.Data;
using Kingdee.BOS.Contracts;
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

using System.Security.Cryptography;
using Kingdee.BOS.Orm.DataEntity;
using Kingdee.BOS.Core;
using Kingdee.BOS.Core.Report;
using Kingdee.BOS.Model.ReportFilter;
using Kingdee.BOS.Core.Bill;
using Kingdee.BOS.Core.DynamicForm.PlugIn;
using Kingdee.BOS.Core.DynamicForm.PlugIn.Args;
using Kingdee.BOS.Core.Metadata.FormElement;
using Kingdee.BOS.BusinessEntity;
using Kingdee.K3.Core.MFG.ENG.ParamOption;
using Kingdee.K3.Core.MFG.ENG.BomExpand;
using Kingdee.K3.MFG.ServiceHelper.ENG;
using Newtonsoft.Json.Linq;
using Kingdee.K3.Core.SCM.STK;
using Kingdee.K3.SCM.ServiceHelper;
using Kingdee.BOS.Core.Metadata.EntityElement;
using System.Net.Mail;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Dynamic;
using ExpressionEvaluator;
using Kingdee.K3.FIN.Core.Parameters;
using Kingdee.K3.FIN.ServiceHelper;
using Kingdee.K3.FIN.Core.Match.Object;
using Kingdee.BOS.Core.CommonFilter;
using Kingdee.BOS.Model.CommonFilter;
using System.Web;
using Kingdee.BOS.Authentication;
using System.Web.Caching;
using Kingdee.BOS.App.Core.ScheduleService;
using System.Transactions;
using Kingdee.BOS.Workflow.App.Core.Repositories;

using Kingdee.BOS.Workflow.Hosting;
using Kingdee.BOS.Workflow.Models;
using Kingdee.BOS.Workflow.Engine;
using Kingdee.BOS.Resource;
using Kingdee.BOS.Workflow.Enums;
using Kingdee.BOS.App.Core;
using Kingdee.BOS.KDThread;
using Kingdee.BOS.Log;
using Kingdee.BOS.Core.Schedules;
using Kingdee.BOS.Workflow.Assignment;
using Kingdee.BOS.Workflow.Models.EnumStatus;
using Kingdee.BOS.Workflow.App.Core;

namespace YOL.K3CLOUD.CUSTOMERNEED
{
    [Kingdee.BOS.Util.HotUpdate]
    public class BillTools
    {
        /// <summary>
        /// 加载单据
        /// </summary>
        /// <param name="billkey">单据标识</param>
        /// <param name="id">单据内码</param>
        /// <param name="ctx">上下文</param>
        /// <returns></returns>
        public Kingdee.BOS.Orm.DataEntity.DynamicObject Loadbill(string billkey, int id, Context ctx)
        {
            IMetaDataService metadataService = Kingdee.BOS.App.ServiceHelper.GetService<IMetaDataService>();
            //获取保存，加载单据服务服务
            //  ISaveService saveService = Kingdee.BOS.App.ServiceHelper.GetService<ISaveService>();
            IViewService viewService = Kingdee.BOS.App.ServiceHelper.GetService<IViewService>();
            //获取元数据
            FormMetadata materialMetadata = metadataService.Load(ctx, billkey) as FormMetadata;
            Kingdee.BOS.Orm.DataEntity.DynamicObject[] objs = viewService.Load(
                                  ctx,
                                  new object[] { id },
                                  materialMetadata.BusinessInfo.GetDynamicObjectType());
            return objs[0];


        }
        public Kingdee.BOS.Orm.DataEntity.DynamicObject Loadbill(string billkey, string id, Context ctx)
        {
            IMetaDataService metadataService = Kingdee.BOS.App.ServiceHelper.GetService<IMetaDataService>();
            //获取保存，加载单据服务服务
            //  ISaveService saveService = Kingdee.BOS.App.ServiceHelper.GetService<ISaveService>();
            IViewService viewService = Kingdee.BOS.App.ServiceHelper.GetService<IViewService>();
            //获取元数据
            FormMetadata materialMetadata = metadataService.Load(ctx, billkey) as FormMetadata;
            Kingdee.BOS.Orm.DataEntity.DynamicObject[] objs = viewService.Load(
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
            Kingdee.BOS.Orm.DataEntity.DynamicObject[] momata = viewService.Load(
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
        public IOperationResult Savebilldata(string billkey, Kingdee.BOS.Orm.DataEntity.DynamicObject data, Context ctx)
        {
            IMetaDataService metadataService = Kingdee.BOS.App.ServiceHelper.GetService<IMetaDataService>();
            //获取保存，加载单据服务服务
            ISaveService saveService = Kingdee.BOS.App.ServiceHelper.GetService<ISaveService>();

            FormMetadata materialMetadata = metadataService.Load(ctx, billkey) as FormMetadata;
            Kingdee.BOS.Orm.DataEntity.DynamicObject[] datas = new Kingdee.BOS.Orm.DataEntity.DynamicObject[] { data };
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
        public Kingdee.BOS.Orm.DataEntity.DynamicObject Getobj(string billkey, Context ctx)
        {
            IMetaDataService metaProxy = Kingdee.BOS.App.ServiceHelper.GetService<IMetaDataService>();
            Kingdee.BOS.Core.Metadata.FormMetadata attachmentMetadata = metaProxy.Load(ctx, billkey) as Kingdee.BOS.Core.Metadata.FormMetadata;
            BusinessInfo attachmentInfo = attachmentMetadata.BusinessInfo;
            Kingdee.BOS.Orm.DataEntity.DynamicObject attachmentData = new Kingdee.BOS.Orm.DataEntity.DynamicObject(attachmentInfo.GetDynamicObjectType());
            return attachmentData;
        }
        /// <summary>
        /// 获取单据体行数据包
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="entitykey"></param>
        /// <returns></returns>
        public Kingdee.BOS.Orm.DataEntity.DynamicObject Getentityobj(Kingdee.BOS.Orm.DataEntity.DynamicObject obj, string entitykey)
        {
            DynamicObjectCollection sktjenys = obj[entitykey] as DynamicObjectCollection;
            Kingdee.BOS.Orm.DataEntity.DynamicObject row = new Kingdee.BOS.Orm.DataEntity.DynamicObject(sktjenys.DynamicCollectionItemPropertyType);
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
        public void DoPush(string sourceFormId, string targetFormId, List<Kingdee.BOS.Orm.DataEntity.DynamicObject> fids, Context ctx, string rulekey, string fbilltype)
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
            Kingdee.BOS.Orm.DataEntity.DynamicObject[] targetBillObjs = (from p in operationResult.TargetDataEntities select p.DataEntity).ToArray();
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
                    Kingdee.BOS.Orm.DataEntity.DynamicObject[] targetobj = new Kingdee.BOS.Orm.DataEntity.DynamicObject[] { targetBillObj };
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
        public void Save(Context ctx, Kingdee.BOS.Orm.DataEntity.DynamicObject obj)
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
        public void SetFZZLDataValue(Kingdee.BOS.Orm.DataEntity.DynamicObject data, BaseDataField bdfield, ref Kingdee.BOS.Orm.DataEntity.DynamicObject dyValue, Context context, string FzZlId)
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
        public void SetBaseDataValue(BaseDataField fmate, long fid, Kingdee.BOS.Orm.DataEntity.DynamicObject data, Context ctx)
        {


            IViewService viewService = ServiceHelper.GetService<IViewService>();
            Kingdee.BOS.Orm.DataEntity.DynamicObject[] stockObjs = viewService.LoadFromCache(ctx,
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
            Message msg = new Kingdee.BOS.Orm.DataEntity.DynamicObject(Message.MessageDynamicObjectType);
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
        public DataSet Connectsqlserver(string serverip, string database, string userid, string password, string sql)
        {
            SqlConnection opensqlserver = new SqlConnection(string.Format("Server={0};Database={1};uid={2};pwd={3}", serverip, database, userid, password, sql));
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
        public string DownloadFile(string url, string download)
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
        public string CreateFileName(string url)
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
        public string Upload(byte[] fileByte, FileInfo file, Context ctx, string loginurl)
        {
            bool bResult = false;
            string sUrl = string.Format("{0}FileUpLoadServices/FileService.svc/Upload2Attachment/?fileName={1}&fileid={2}&token={5}&dbid={3}&last={4}",
                   loginurl + "/", file.Name, string.Empty, ctx.DBId, true, ctx.UserToken);
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
                //System.IO.File.Move(tempFile, path);
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
            string fileExt = url.Substring(url.LastIndexOf(@"/")).Replace("/", "");
            fileName = fileExt;
            return fileName;
        }
        /// <summary>
        /// md5加密转大写 str需加密字符串，code 16位OR32位置
        /// </summary>
        /// <param name="str"></param>
        /// <param name="code"></param>
        /// <returns></returns>
        public string md5(string str, int code)
        {
            if (code == 16) //16位MD5加密（取32位加密的9~25字符） 
            {
                return System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(str, "MD5").ToUpper().Substring(8, 16);
            }
            else//32位加密 
            {
                return System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(str, "MD5").ToUpper();
            }
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
        public string PYAESEncrypt(string text, string password, string iv)
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
                   new object[] { id });
            return result.IsSuccess;
        }
        /// <summary>
        /// 获取到期日，ywdate业务日期，sktjeny收款条件分录数据包
        /// </summary>
        /// <param name="ywdate"></param>
        /// <param name="sktjeny"></param>
        /// <returns></returns>
        public DateTime GetDQR(DateTime ywdate, Kingdee.BOS.Orm.DataEntity.DynamicObject sktjeny)//根据收款条件计算到期日,应收单日期,收款计划单据体数据包
        {

            if (sktjeny != null)
            {

                DateTime xtdqr = ywdate;


                bool yd = Convert.ToBoolean(sktjeny["ISMONTHEND"]);//是否勾选月底
                Kingdee.BOS.Orm.DataEntity.DynamicObject dqrfs = sktjeny["DuecalMethodID"] as Kingdee.BOS.Orm.DataEntity.DynamicObject;//到期日结算方式
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

        /// <summary>
        /// 出入库单据根据物料，仓库，批号查询物料收发汇总表（用于校验先出后进库存单据），返回临时表集合 bill 单据数据包，fentitykey单据体实体名，zzid库存组织内码，famorm物料实体名，fstockorm仓库实体名，fbaseqtyorm 基本数量实体名，ctx上下文
        /// </summary>
        /// <param name="bill"></param>
        /// <param name="fentityorm"></param>
        /// <param name="zzid"></param>
        /// <param name="fmaorm"></param>
        /// <param name="fstockorm"></param>
        /// <param name="flotorm"></param>
        /// <param name="fbaseqtyrom"></param>
        /// <param name="ctx"></param>
        /// <returns></returns>
        public List<DataTable> Temptables(Kingdee.BOS.Orm.DataEntity.DynamicObject bill, string fentityorm, string zzid, string fmaorm, string fstockorm, string flotorm, string fbaseqtyorm, Context ctx, string dateorm)
        {
            List<DataTable> temptable = new List<DataTable>();
            if (fmaorm == "MaterialIDSETY") //组装拆卸单
            {
                DynamicObjectCollection zzenys = bill[fentityorm] as DynamicObjectCollection;
                foreach (var zzeny in zzenys)
                {
                    DynamicObjectCollection enys = zzeny["STK_ASSEMBLYSUBITEM"] as DynamicObjectCollection;
                    foreach (var eny in enys)
                    {
                        decimal ckqty = Convert.ToDecimal(eny[fbaseqtyorm]);
                        Kingdee.BOS.Contracts.ISysReportService sysReporSservice = ServiceFactory.GetSysReportService(ctx);
                        Kingdee.BOS.Contracts.IPermissionService permissionService = ServiceFactory.GetPermissionService(ctx);
                        var filterMetadata = FormMetaDataCache.GetCachedFilterMetaData(ctx);//加载字段比较条件元数据。
                        var reportMetadata = FormMetaDataCache.GetCachedFormMetaData(ctx, "STK_StockSummaryRpt");//
                        var reportFilterMetadata = FormMetaDataCache.GetCachedFormMetaData(ctx, "STK_StockSummaryFilter");//
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
                        var parameterData = UserParamterServiceHelper.Load(ctx, parameterDataMetadata.BusinessInfo, ctx.UserId, "STK_StockSummaryRpt", KeyConst.USERPARAMETER_KEY);
                        //  var entity = model.Load("3bf9ec9456714af5a2b73658b5221dca");//过滤方案的主键值，可通过该SQL语句查询得到：SELECT * FROM T_BAS_FILTERSCHEME
                        var filter = model.GetFilterParameter();
                        Kingdee.BOS.Orm.DataEntity.DynamicObject flter = filter.CustomFilter;
                        //   flter["StockOrgId"] = bill["PickOrgId"];
                        flter["StockOrgId"] = zzid;

                        //    List<int> orgs = new List<int>();
                        //    orgs.Add(Convert.ToInt32(this.Context.CurrentOrganizationInfo.ID));
                        //    datas.SetXLLBValue(flter, orgs, "StockOrgId", Convert.ToInt32(this.Context.CurrentOrganizationInfo.ID));
                        flter["BeginDate"] = DateTime.Now;
                        flter["EndDate"] = Convert.ToDateTime(bill[dateorm]).AddMonths(+1).AddDays(-Convert.ToDateTime(bill[dateorm]).Day);
                        //  flter["BeginMaterialId_Id"] = eny["MaterialId_Id"];
                        flter["BeginMaterialId"] = eny[fmaorm];
                        //    flter["BeginStockId_Id"] = eny["StockId_Id"];
                        flter["BeginStockId"] = eny[fstockorm];
                        flter["BeginLotNo"] = eny[flotorm];
                        flter["BeginLotNo_Text"] = eny[flotorm + "_Text"];
                        //     flter["BeginLotNo_Id"] = eny["Lot_Id"];
                        //     flter["EndMaterialId_Id"] = eny["MaterialId_Id"];
                        flter["EndMaterialId"] = eny[fmaorm];
                        //      flter["EndStockId_Id"] = eny["StockId_Id"];
                        flter["EndStockId"] = eny[fstockorm];
                        flter["EndLotNo"] = eny[flotorm];
                        flter["EndLotNo_Text"] = eny[flotorm + "_Text"];
                        //   flter["EndLotNo_Id"] = eny["Lot_Id"];

                        flter["BillStatus"] = "C";
                        flter["QCPriceSource"] = "3";
                        flter["IOPriceSource"] = "1";
                     //   flter["TransDirectPriceSource"] = "2";//7.3没有
                     //   flter["VMIInstockPriceSource"] = "2";

                      //  flter["FIsHsStockField"] = false;
                        flter["IncludReceive"] = false;
                        flter["ShowForbidMaterial"] = false;
                        flter["NoHappenNoShow"] = false;
                        flter["SplitPageByOwner"] = true;
                        flter["NoQtyNoShow"] = false;
                        flter["NoBalanceNoShow"] = false;
                        flter["NoStockAdjust"] = false;
                        flter["NoHappenNoQcNoShow"] = false;
                        flter["NoInnerTrans"] = false;
                      // flter["StockQtyRule"] = "0";
                        IRptParams p = new RptParams();
                        p.FormId = reportFilterMetadata.BusinessInfo.GetForm().Id;
                        p.StartRow = 1;
                        p.EndRow = int.MaxValue;//StartRow和EndRow是报表数据分页的起始行数和截至行数，一般取所有数据，所以EndRow取int最大值。
                        p.FilterParameter = filter;
                        p.FilterFieldInfo = model.FilterFieldInfo;
                        p.BaseDataTempTable.AddRange(permissionService.GetBaseDataTempTable(ctx, reportMetadata.BusinessInfo.GetForm().Id));
                        p.ParameterData = parameterData;
                        p.CustomParams.Add(KeyConst.OPENPARAMETER_KEY, openParameter);
                        using (DataTable dt = sysReporSservice.GetData(ctx, reportMetadata.BusinessInfo, p))
                        {
                            if (!(dt != null && dt.Rows.Count > 0))
                            {
                                DataTable tablez = new DataTable();
                                tablez.Columns.Add("FSEQ");
                                tablez.Columns.Add("FCKQTY");
                                tablez.Columns.Add("FBASEJCQTY");
                                DataRow row = tablez.NewRow();
                                row["FSEQ"] = eny["SEQ"].ToString();
                                row["FBASEJCQTY"] = Convert.ToDecimal(0);
                                row["FCKQTY"] = Convert.ToDecimal(eny[fbaseqtyorm]);
                                tablez.Rows.Add(row);
                                temptable.Add(tablez);


                            }
                            else
                            {
                                dt.Columns.Add("FSEQ");
                                dt.Columns.Add("FCKQTY");
                                dt.Rows[0]["FSEQ"] = eny["SEQ"].ToString();
                                dt.Rows[0]["FCKQTY"] = Convert.ToDecimal(eny[fbaseqtyorm]);
                                temptable.Add(dt);
                            }
                        }
                        ServiceFactory.CloseService(sysReporSservice);
                        ServiceFactory.CloseService(permissionService);
                    }


                }
            }
            else
            {
                DynamicObjectCollection enys = bill[fentityorm] as DynamicObjectCollection;
                foreach (var eny in enys)
                {
                    decimal ckqty = Convert.ToDecimal(eny[fbaseqtyorm]);
                    Kingdee.BOS.Contracts.ISysReportService sysReporSservice = ServiceFactory.GetSysReportService(ctx);
                    Kingdee.BOS.Contracts.IPermissionService permissionService = ServiceFactory.GetPermissionService(ctx);
                    var filterMetadata = FormMetaDataCache.GetCachedFilterMetaData(ctx);//加载字段比较条件元数据。
                    var reportMetadata = FormMetaDataCache.GetCachedFormMetaData(ctx, "STK_StockSummaryRpt");//
                    var reportFilterMetadata = FormMetaDataCache.GetCachedFormMetaData(ctx, "STK_StockSummaryFilter");//
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
                    var parameterData = UserParamterServiceHelper.Load(ctx, parameterDataMetadata.BusinessInfo, ctx.UserId, "STK_StockSummaryRpt", KeyConst.USERPARAMETER_KEY);
                    //  var entity = model.Load("3bf9ec9456714af5a2b73658b5221dca");//过滤方案的主键值，可通过该SQL语句查询得到：SELECT * FROM T_BAS_FILTERSCHEME
                    var filter = model.GetFilterParameter();
                    Kingdee.BOS.Orm.DataEntity.DynamicObject flter = filter.CustomFilter;
                    //   flter["StockOrgId"] = bill["PickOrgId"];
                    flter["StockOrgId"] = zzid;

                    //    List<int> orgs = new List<int>();
                    //    orgs.Add(Convert.ToInt32(this.Context.CurrentOrganizationInfo.ID));
                    //    datas.SetXLLBValue(flter, orgs, "StockOrgId", Convert.ToInt32(this.Context.CurrentOrganizationInfo.ID));
                    flter["BeginDate"] = DateTime.Now;
                    flter["EndDate"] = Convert.ToDateTime(bill[dateorm]).AddMonths(+1).AddDays(-Convert.ToDateTime(bill[dateorm]).Day);
                    //  flter["BeginMaterialId_Id"] = eny["MaterialId_Id"];
                    flter["BeginMaterialId"] = eny[fmaorm];
                    //    flter["BeginStockId_Id"] = eny["StockId_Id"];
                    flter["BeginStockId"] = eny[fstockorm];
                    flter["BeginLotNo"] = eny[flotorm];
                    flter["BeginLotNo_Text"] = eny[flotorm + "_Text"];
                    //     flter["BeginLotNo_Id"] = eny["Lot_Id"];
                    //     flter["EndMaterialId_Id"] = eny["MaterialId_Id"];
                    flter["EndMaterialId"] = eny[fmaorm];
                    //      flter["EndStockId_Id"] = eny["StockId_Id"];
                    flter["EndStockId"] = eny[fstockorm];
                    flter["EndLotNo"] = eny[flotorm];
                    flter["EndLotNo_Text"] = eny[flotorm + "_Text"];
                    //   flter["EndLotNo_Id"] = eny["Lot_Id"];

                    flter["BillStatus"] = "C";
                    flter["QCPriceSource"] = "3";
                    flter["IOPriceSource"] = "1";
                 //   flter["TransDirectPriceSource"] = "2";//7.3版本没有此字段
                 //   flter["VMIInstockPriceSource"] = "2";

                 //   flter["FIsHsStockField"] = false;
                    flter["IncludReceive"] = false;
                    flter["ShowForbidMaterial"] = false;
                    flter["NoHappenNoShow"] = false;
                    flter["SplitPageByOwner"] = true;
                    flter["NoQtyNoShow"] = false;
                    flter["NoBalanceNoShow"] = false;
                    flter["NoStockAdjust"] = false;
                    flter["NoHappenNoQcNoShow"] = false;
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
                    using (DataTable dt = sysReporSservice.GetData(ctx, reportMetadata.BusinessInfo, p))
                    {
                        if (!(dt != null && dt.Rows.Count > 0))
                        {
                            DataTable tablez = new DataTable();
                            tablez.Columns.Add("FSEQ");
                            tablez.Columns.Add("FCKQTY");
                            tablez.Columns.Add("FBASEJCQTY");
                            DataRow row = tablez.NewRow();
                            row["FSEQ"] = eny["SEQ"].ToString();
                            row["FBASEJCQTY"] = Convert.ToDecimal(0);
                            row["FCKQTY"] = Convert.ToDecimal(eny[fbaseqtyorm]);
                            tablez.Rows.Add(row);
                            temptable.Add(tablez);


                        }
                        else
                        {
                            dt.Columns.Add("FSEQ");
                            dt.Columns.Add("FCKQTY");
                            dt.Rows[0]["FSEQ"] = eny["SEQ"].ToString();
                            dt.Rows[0]["FCKQTY"] = Convert.ToDecimal(eny[fbaseqtyorm]);
                            temptable.Add(dt);
                        }
                    }
                    ServiceFactory.CloseService(sysReporSservice);
                    ServiceFactory.CloseService(permissionService);


                }
            }
            return temptable;
        }
        /// <summary>
        /// 操作插件调用字段值更新，ctx上下文，billid单据内码，billkey单据标识，fieldnewvalue字段新值，fieldkey字段标识,单据头字段row传负数，单据体字段传对应行
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="billid"></param>
        /// <param name="billkey"></param>
        /// <param name="fieldnewvalue"></param>
        /// <param name="fieldkey"></param>
        public void Datachange(Context ctx, string billid, string billkey, string fieldnewvalue, string fieldkey, int row)
        {
            if (billid == null)
            {
                return;
            }
            else
            {
                IBillView billView = this.CreateBillView(ctx, billkey, billid);
                if (row < 0)
                {
                    billView.Model.SetValue(fieldkey, 0);
                    billView.Model.SetValue(fieldkey, fieldnewvalue);

                }
                else
                {
                    billView.Model.SetValue(fieldkey, 0, row);
                    billView.Model.SetValue(fieldkey, fieldnewvalue, row);

                }
                Kingdee.BOS.ServiceHelper.BusinessDataServiceHelper.Save(ctx, billView.Model.DataObject);

            }


        }
        /// <summary>
        /// 创建view模拟人为操作单据，ctx上下文，key单据标识，id单据内码
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="key"></param>
        /// <param name="id"></param>
        /// <returns></returns>
        public IBillView CreateBillView(Context ctx, string key, string id)
        {
            // 读取物料的元数据
            FormMetadata meta = MetaDataServiceHelper.Load(ctx, key) as FormMetadata;
            Form form = meta.BusinessInfo.GetForm();
            // 创建用于引入数据的单据view
            Type type = Type.GetType("Kingdee.BOS.Web.Import.ImportBillView,Kingdee.BOS.Web");
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
        /// <summary>
        /// 发送邮件，ctx当前上下文，userid发送人内码，title 标题，message 内容 ，adress 接收邮箱地址，需要先配置好发件人的邮箱设置
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="userid"></param>
        /// <param name="title"></param>
        /// <param name="message"></param>
        /// <param name="adress"></param>
        /// <returns></returns>
        public ResultInfo resultInfo(Context ctx, long userid, string title, string message, string[] adress)
        {
            ResultInfo resultInfo = SendMailServiceHelper.Send(ctx, userid, title, message, adress);
            return resultInfo;
        }

        /// <summary>
        /// /物料清单正查，fmaid物料内码，fbom物料清单内码，fbaseunitid物料基本单位内码
        /// </summary>
        /// <param name="fmaid"></param>
        /// <param name="fbomid"></param>
        /// <param name="fbaseunitid"></param>
        /// <returns></returns>
        public DynamicObjectCollection bomenys(int fmaid, int fbomid, int fbaseunitid, Context ctx)
        {
            MemBomExpandOption bomQueryOption = new MemBomExpandOption();
            bomQueryOption.ExpandLevelTo = 100;
            bomQueryOption.DeleteVirtualMaterial = false;
            bomQueryOption.ExpandVirtualMaterial = true;
            bomQueryOption.DeleteSkipRow = false;
            bomQueryOption.ExpandSkipRow = true;
            bomQueryOption.CsdSubstitution = true;
            bomQueryOption.IsExpandSupplyManager = true;
            bomQueryOption.ParentCsdYieldRate = true;
            bomQueryOption.ChildCsdYieldRate = true;

            List<Kingdee.BOS.Orm.DataEntity.DynamicObject> lstRows = new List<Kingdee.BOS.Orm.DataEntity.DynamicObject>();
            BomForwardSourceDynamicRow row = BomForwardSourceDynamicRow.CreateInstance();
            row.MaterialId_Id = fmaid;
            row.BomId_Id = fbomid;
            row.BaseUnitId_Id = fbaseunitid;
            row.NeedQty = 1;
            row.NeedDate = DateTime.Now;
            row.WorkCalId_Id = 0;
            row.TimeUnit = 1.ToString();

            lstRows.Add(row);
            Kingdee.BOS.Orm.DataEntity.DynamicObject bomExpandData = BomExpandServiceHelper.ExpandBomForwardMen(ctx, lstRows, bomQueryOption);
            DynamicObjectCollection results = bomExpandData["BomExpandResult"] as DynamicObjectCollection;
            return results;
        }
        /// <summary>
        /// 调用操作事件，billkey单据标识，pkIds（fid 单据内码，fentrydid 分录内码 调用行操作时要传入 ），czid操作编码，ctx上下文
        /// </summary>
        /// <param name="billkey"></param>
        /// <param name="pkIds"></param>
        /// <param name="czid"></param>
        /// <param name="ctx"></param>
        /// <returns></returns>
        public IOperationResult dycz(string billkey, List<KeyValuePair<object, object>> pkIds, string czid, Context ctx)
        {
            IMetaDataService metadataServicez = Kingdee.BOS.App.ServiceHelper.GetService<IMetaDataService>();
            FormMetadata materialMetadataz = metadataServicez.Load(ctx, billkey) as FormMetadata;
            var result = BusinessDataServiceHelper.SetBillStatus(ctx, materialMetadataz.BusinessInfo, pkIds, null, czid, OperateOption.Create());
            return result;
        }
        /// <summary>
        /// 按月份汇总查询物料收发汇总表，ctx上下文，zzid组织，yf月份，year 年份 ,billstate单据状态
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="zzid"></param>
        /// <param name="yf"></param>
        /// <param name="year"></param>
        /// <param name="billstate"></param>
        /// <returns></returns>
        public string temptable(Context ctx, string zzid, string yf, string year, string billstate,string gjgl)
        {

          //  decimal sumqty = 0;
            DateTime time = Convert.ToDateTime(year +"-"+ yf + "-01");
            ISysReportService sysReporSservice = ServiceFactory.GetSysReportService(ctx);
            IPermissionService permissionService = ServiceFactory.GetPermissionService(ctx);
            var filterMetadata = FormMetaDataCache.GetCachedFilterMetaData(ctx);//加载字段比较条件元数据。
            var reportMetadata = FormMetaDataCache.GetCachedFormMetaData(ctx, "STK_StockSummaryRpt");//
            var reportFilterMetadata = FormMetaDataCache.GetCachedFormMetaData(ctx, "STK_StockSummaryFilter");//
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
            var parameterData = UserParamterServiceHelper.Load(ctx, parameterDataMetadata.BusinessInfo, ctx.UserId, "STK_StockSummaryRpt", KeyConst.USERPARAMETER_KEY);

            var filter = model.GetFilterParameter();
            Kingdee.BOS.Orm.DataEntity.DynamicObject flter = filter.CustomFilter;
            flter["StockOrgId"] = zzid;
            flter["BeginDate"] = time;
            flter["EndDate"] = time.AddMonths(+1).AddDays(-1);
            ////  flter["BeginMaterialId_Id"] = eny["MaterialId_Id"];
            //flter["BeginMaterialId"] = eny[fmaorm];
            ////    flter["BeginStockId_Id"] = eny["StockId_Id"];
            //flter["BeginStockId"] = eny[fstockorm];
            //flter["BeginLotNo"] = eny[flotorm];
            //flter["BeginLotNo_Text"] = eny[flotorm + "_Text"];
            //     flter["BeginLotNo_Id"] = eny["Lot_Id"];
            //     flter["EndMaterialId_Id"] = eny["MaterialId_Id"];
            //flter["EndMaterialId"] = eny[fmaorm];
            ////      flter["EndStockId_Id"] = eny["StockId_Id"];
            //flter["EndStockId"] = eny[fstockorm];
            //flter["EndLotNo"] = eny[flotorm];
            //flter["EndLotNo_Text"] = eny[flotorm + "_Text"];
            //   flter["EndLotNo_Id"] = eny["Lot_Id"];
            flter["BillStatus"] = billstate;
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
            if (gjgl!=null)
            {
                p.FilterParameter.FilterString = gjgl;
            }
           
            p.FilterFieldInfo = model.FilterFieldInfo;
            p.BaseDataTempTable.AddRange(permissionService.GetBaseDataTempTable(ctx, reportMetadata.BusinessInfo.GetForm().Id));
            p.ParameterData = parameterData;
            p.CustomParams.Add(KeyConst.OPENPARAMETER_KEY, openParameter);
            MoveReportServiceParameter report = new MoveReportServiceParameter(ctx, reportMetadata.BusinessInfo, Guid.NewGuid().ToString(),p);
           
            var aa = sysReporSservice.GetListAndReportData(report);
            DataTable dt = aa.DataSource;
            ServiceFactory.CloseService(sysReporSservice);
            ServiceFactory.CloseService(permissionService);
            if (dt!=null)
            {
                
                return dt.TableName;

                //if (dt.Compute("sum(FBASEOUTQTY)", "").IsNullOrEmptyOrWhiteSpace())
                //{
                //    sumqty = 0;
                //}
                //else
                //{
                //    sumqty = Convert.ToDecimal(dt.Compute("sum(FBASEOUTQTY)", ""));
                //}
            }
            else
            {
                return null;
            }
           
           
            //string tablename=  sysReporSservice.GetDataTableName(ctx, reportMetadata.BusinessInfo, p);
            //DataTable table = this.Select(string.Format("select isnull(sum(FBASEOUTQTY),0) FBASEOUTQTY from {0}", tablename), ctx);
            //if (table!=null)
            //{
            //    sumqty = Convert.ToDecimal(table.Rows[0]["FBASEOUTQTY"]);
            //}
            //else
            //{
            //    sumqty = 0;
            //}
            //using (DataTable dt = sysReporSservice.GetData(ctx, reportMetadata.BusinessInfo, p,true))
            //{

            //    if (dt.Compute("sum(FBASEOUTQTY)", "").IsNullOrEmptyOrWhiteSpace())
            //    {
            //        sumqty = 0;
            //    }
            //    else
            //    {
            //        sumqty = Convert.ToDecimal(dt.Compute("sum(FBASEOUTQTY)", ""));
            //    }


            //}
          

            


        }

        /// <summary>
        /// 此方法模拟人为查询物料收发汇总表数据返回DataTable拿到DataTable自行操作DataTable获取需要数据，ctx上下文，zzid 组织内码，yf月份，year 年份，billstate单据状态,fmadata物料数据包,fstockdata仓库数据包
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="zzid"></param>
        /// <param name="yf"></param>
        /// <param name="year"></param>
        /// <param name="flot"></param>
        /// <param name="fmadata"></param>
        /// <param name="fstockdata"></param>
        /// <returns></returns>
        public DataTable Wlsfhzb(Context ctx, string zzid, string starttime, string endtime, int fmaid, int fstockid, string flot, string billstate, Kingdee.BOS.Orm.DataEntity.DynamicObject fmadata, Kingdee.BOS.Orm.DataEntity.DynamicObject fstockdata)
        {
            DataTable temptable = null;

            ISysReportService sysReporSservice = ServiceFactory.GetSysReportService(ctx);
            IPermissionService permissionService = ServiceFactory.GetPermissionService(ctx);
            var filterMetadata = FormMetaDataCache.GetCachedFilterMetaData(ctx);//加载字段比较条件元数据。
            var reportMetadata = FormMetaDataCache.GetCachedFormMetaData(ctx, "STK_StockSummaryRpt");//物料收发汇总表标识
            var reportFilterMetadata = FormMetaDataCache.GetCachedFormMetaData(ctx, "STK_StockSummaryFilter");//物料收发汇总表过滤框
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
            var parameterData = UserParamterServiceHelper.Load(ctx, parameterDataMetadata.BusinessInfo, ctx.UserId, "STK_StockSummaryRpt", KeyConst.USERPARAMETER_KEY);

            var filter = model.GetFilterParameter();
            Kingdee.BOS.Orm.DataEntity.DynamicObject flter = filter.CustomFilter;
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

        /// <summary>
        /// 金蝶云发送邮件方法 ctx上下文，fid单据内码，billkey单据主键，billtypeid单据类型id，mailadress收件人地址，title邮件标题，msg邮件内容，yjdata邮件账户设置(BAS_MAILACCOUNTSETTING)数据包，pdf是否发送pdf,excel是否发送excel
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="fid"></param>
        /// <param name="billkey"></param>
        /// <param name="billtypeid"></param>
        /// <param name="mailadress"></param>
        /// <param name="title"></param>
        /// <param name="msg"></param>
        /// <param name="yjdata"></param>
        /// <param name="pdf"></param>
        /// <param name="excel"></param>
        /// <returns></returns>
        public string sendmail(Context ctx, string fid, string billkey, string billtypeid, string[] mailadress, string title, string msg, Kingdee.BOS.Orm.DataEntity.DynamicObject yjdata, bool pdf, bool excel)
        {
            if (yjdata != null)
            {
                Kingdee.BOS.Orm.DataEntity.DynamicObject yjdzdata = yjdata["FSENDMAILSERVER"] as Kingdee.BOS.Orm.DataEntity.DynamicObject;
                string yjdz = null;//邮件地址
                int yjdk = 0;//邮件端口
                string fwdz = null;//邮件服务器地址
                string zh = null;//发送邮件账户
                string sqm = null;//授权码
                string detdmb = null;//默认套打模板
                if (!yjdata["FEmailAddress"].IsNullOrEmptyOrWhiteSpace())
                {
                    fwdz = yjdata["FEmailAddress"].ToString();
                }
                if (!yjdata["FUserName"].IsNullOrEmptyOrWhiteSpace())
                {
                    zh = yjdata["FUserName"].ToString();
                }
                if (!yjdata["FPassword"].IsNullOrEmptyOrWhiteSpace())
                {
                    sqm = yjdata["FPassword"].ToString();
                }
                if (yjdzdata != null)
                {
                    yjdz = yjdzdata["FMailServer"].ToString();
                    yjdk = Convert.ToInt32(yjdzdata["FSMTPPort"]);
                }

                IDynamicFormView view = this.CreateBillView(ctx, billkey, fid);

                string paratext = UserParamterServiceHelper.Load(ctx, "NotePrintSetup" + view.BillBusinessInfo.GetForm().Id.ToUpper().GetHashCode().ToString(), ctx.UserId, "");//获取套打设置参数

                JArray paras = JArray.Parse(paratext);
                if (paras != null)
                {
                    var tdmbs = (from pa in paras where pa["key"].ToString() == billtypeid select pa).ToList();
                    if (tdmbs.Count > 0)
                    {
                        detdmb = tdmbs[0]["value"].ToString();
                    }
                }
                List<string> bids = new List<string>();
                bids.Add(fid);
                List<string> tdids = new List<string>();
                tdids.Add(detdmb);
                //(view as IDynamicFormViewService).MainBarItemClick("YOL_YJ");
                Kingdee.BOS.Core.Import.IImportView importView = view as Kingdee.BOS.Core.Import.IImportView;

                importView.AddViewSession();

                //PrintExportInfo pExInfo = new PrintExportInfo();

                //pExInfo.PageId = view.PageId;

                //pExInfo.FormId = "PUR_Requisition";

                //pExInfo.BillIds = bids;//单据内码

                //pExInfo.TemplateIds = tdids;//套打模板ID

                //pExInfo.FileType = ExportFileType.PDF;//文件格式

                //pExInfo.ExportType = ExportType.ByPage;//导出格式

                //string temppath = @"E:\K3FILE";
                //string filePath = Path.Combine(temppath, Guid.NewGuid().ToString() + ".PDF");
                //pExInfo.FilePath = filePath;//文件输出路径

                //IDynamicFormViewService yjview = view as IDynamicFormViewService;

                //yjview.ExportNotePrint(pExInfo);

                //  importView.RemoveViewSession();
                string[] ids = { fid };
                string[] cc = null;
                string exportFileName = null;
                string pdfExportFileName = null;
                string excelExportFileName = null;
                ResultInfo resultInfo = null;
                DynamicFormViewPlugInProxy service = view.GetService<DynamicFormViewPlugInProxy>();
                if (service != null)
                {
                    InitializeSendMailServiceEventArgs initializeSendMailServiceEventArgs = new InitializeSendMailServiceEventArgs();
                    service.FireOnInitializeSendMailService(initializeSendMailServiceEventArgs);
                    exportFileName = initializeSendMailServiceEventArgs.ExportFileName;

                }
                try
                {
                    pdfExportFileName = SendMailServiceHelper.GetPdfExportFileName(view, ctx.UserId, ids, detdmb, exportFileName);
                }
                catch (Exception)
                {

                    return "生成pdf失败";
                }
                try
                {
                    excelExportFileName = SendMailServiceHelper.GetExcelExportFileName(view, ctx.UserId, ids, detdmb, exportFileName);
                }
                catch (Exception)
                {

                    return "生成excel失败";
                }
                try
                {
                    resultInfo = SendMailServiceHelper.Send(view, ctx.UserId, yjdz, yjdk, false, false, fwdz, zh, sqm, title, msg, mailadress, cc, ids, pdf, pdfExportFileName, excel, excelExportFileName, exportFileName);
                }
                catch (Exception ex)
                {

                    return ex.Message;
                }
                if (resultInfo.Successful == true)
                {
                    return "邮件发送成功！";
                }
                else
                {
                    return "邮件发送失败！原因：" + resultInfo.Message;
                }


            }
            else
            {
                return "参数不能为空！";
            }
        }

        /// <summary>
        /// 调用存储过程（sqlserver）获取流水号，ctx当前上下文,bmyj编码依据（日期等字段,不同的编码依据创建不同的流水号表）,需现在数据库创建表以及存储过程
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="bmyj"></param>
        /// <returns></returns>
        public string Getrm(Context ctx, string bmyj)
        {
            if (bmyj == null)
            {
                return null;
            }
            //string sqlp = "/*dialect*/  create procedure getrm @fildname varchar(500) ,@seed  int output as declare @fild nvarchar(500),@sql nvarchar(max) set nocount on set XACT_ABORT ON set @fild=@fildname set @sql=' select @seed=isnull(max(id),0) from seeds with(TABLOCKX) where fieldname=@fild'  exec sp_executesql @sql, N'@seed int out,@fild nvarchar(500)',@seed out,@fild if @@rowcount=0 or @seed=0 begin set @sql='insert into seeds(id,fieldname) values(1,@fild)' exec sp_executesql @sql , N'@fild nvarchar(500) ',@fild  set @seed =1 set @sql=' update seeds set id=@seed+1 where fieldname= @fild' exec sp_executesql @sql, N'@seed int out  ,@fild nvarchar(500) ',@seed  out,@fild return @seed end if @seed<>0 begin set @sql=' update seeds set id=@seed+1 where fieldname= @fild'exec sp_executesql @sql, N'@seed int out  ,@fild nvarchar(500) ',@seed  out,@fild return @seed end ";
            //int row = DBServiceHelper.Execute(ctx, sqlp);//创建存储过程
            //string sqlt = "/*dialect*/  CREATE TABLE [dbo].[seeds]([id] [int] NOT NULL,[fieldname] [varchar](100) NOT NULL) ON [PRIMARY] ";
            //int rowt = DBServiceHelper.Execute(ctx, sqlt);//创建流水号表
            string sql = string.Format("/*dialect*/  begin tran DECLARE	@return_value int,@seed int EXEC	@return_value = getrm @fildname = '{0}',@seed = @seed OUTPUT SELECT	'ReturnValue' = @return_value commit", bmyj);
            DataSet tables = DBServiceHelper.ExecuteDataSet(ctx, sql);
            if (!(tables != null && tables.Tables.Count > 0 && tables.Tables[0].Rows.Count > 0))
            {
                return null;
            }
            else
            {
                return tables.Tables[0].Rows[0]["ReturnValue"].ToString();
            }
            // create procedure[dbo].[getrm] --产生流水号
            // @fildname varchar(500) ,@seed int output
            // as
            // declare
            //  @fild nvarchar(500),
            //  @sql nvarchar(max)
            //  set nocount on
            //  set XACT_ABORT ON
            //  set @fild = @fildname
            //  set @sql = ' select @seed=isnull(max(id),0) from seeds with(TABLOCKX) where fieldname=@fild'--with(TABLOCKX)设置锁
            //  exec sp_executesql @sql, N'@seed int out,@fild nvarchar(500)', @seed out, @fild 
            //  if @@rowcount=0 or @seed = 0
            //  begin
            //  set @sql='insert into seeds(id,fieldname) values(1,@fild)'
            //  exec sp_executesql @sql , N'@fild nvarchar(500) ',@fild
            //  set @seed =1
            //  set @sql = ' update seeds set id=@seed+1 where fieldname= @fild'
            //  exec sp_executesql @sql, N'@seed int out  ,@fild nvarchar(500) ',@seed  out,@fild
            //  return @seed
            //  end
            //  if @seed<>0
            //  begin
            //  set @sql=' update seeds set id=@seed+1 where fieldname= @fild'
            //  exec sp_executesql @sql, N'@seed int out  ,@fild nvarchar(500) ',@seed  out,@fild
            //  return @seed
            //  end

        }
        /// <summary>
        /// 获取多个流水号，ctx上下文，bmyj编码依据，counts流水号个数
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="bmyj"></param>
        /// <param name="counts"></param>
        /// <returns></returns>
        public List<string> Getrms(Context ctx, string bmyj, int counts)
        {
            if (bmyj == null)
            {
                return null;
            }

            string sql = string.Format("/*dialect*/  begin tran DECLARE	@return_value int,@seed int EXEC	@return_value = getrms @fildname = '{0}',@seed = @seed OUTPUT,@counts = {1} SELECT	'ReturnValue' = @return_value commit", bmyj, counts);
            DataSet tables = DBServiceHelper.ExecuteDataSet(ctx, sql);
            if (!(tables != null && tables.Tables.Count > 0 && tables.Tables[0].Rows.Count > 0))
            {
                return null;
            }
            else
            {
                int startnumber = Convert.ToInt32(tables.Tables[0].Rows[0]["ReturnValue"]);
                List<string> numbers = new List<string>();
                for (int i = 0; i < counts; i++)
                {
                    numbers.Add((startnumber + i).ToString());
                }
                return numbers;
            }


        }
        /// <summary>
        /// 表单插件中获取批号方法，ctx上下文，view this.View,flotkey 批号字段标识,lotorm批号字段绑定实体属性
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="view"></param>
        /// <param name="flotkey"></param>
        /// <param name="lotorm"></param>
        public void GetFlotBill(Context ctx, IBillView view, string flotkey, string lotorm)
        {
            LotField lotField = view.Model.BusinessInfo.GetField(flotkey) as LotField;
            ExtendedDataEntitySet extendedDataEntitySet = new ExtendedDataEntitySet();
            extendedDataEntitySet.Parse(new Kingdee.BOS.Orm.DataEntity.DynamicObject[] { view.Model.DataObject }, view.Model.BusinessInfo);
            ExtendedDataEntity[] array = extendedDataEntitySet.FindByEntityKey(lotField.EntityKey);
            CodeAppResult codeAppResult = StockServiceHelper.GenerateLotMasterByCodeRule(ctx, view.Model.BillBusinessInfo, lotField, array);
            var flots = codeAppResult.CodeNumbers;
            Entity entity = view.Model.BusinessInfo.GetEntryEntity(lotField.EntityKey);
            for (int i = 0; i < flots.Count; i++)
            {
                if (!flots[i].Value.IsNullOrEmptyOrWhiteSpace())
                {
                    Kingdee.BOS.Orm.DataEntity.DynamicObject obj = view.Model.GetEntityDataObject(entity, i);
                    obj[lotorm + "_Text"] = flots[i].Value[0];
                }
            }
            view.UpdateView(lotField.EntityKey);
        }
        /// <summary>
        /// 单据转换插件中获取批号（必须在AfterConvert事件中使用） ctx当前上下文,flotkey批号字段标识，lotorm批号绑定实体属性，enys转换后事件单据体数据包，obj单据数据包
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="billkey"></param>
        /// <param name="flotkey"></param>
        /// <param name="lotorm"></param>
        public void GetFlot(Context ctx, string billkey, string flotkey, string lotorm, DynamicObjectCollection enys, Kingdee.BOS.Orm.DataEntity.DynamicObject obj)
        {

            IMetaDataService metadataService = Kingdee.BOS.App.ServiceHelper.GetService<IMetaDataService>();
            FormMetadata materialMetadata = metadataService.Load(ctx, billkey) as FormMetadata;
            LotField lotField = materialMetadata.BusinessInfo.GetField(flotkey) as LotField;
            ExtendedDataEntitySet extendedDataEntitySet = new ExtendedDataEntitySet();
            extendedDataEntitySet.Parse(new Kingdee.BOS.Orm.DataEntity.DynamicObject[] { obj }, materialMetadata.BusinessInfo);
            ExtendedDataEntity[] array = extendedDataEntitySet.FindByEntityKey(lotField.EntityKey);
            CodeAppResult codeAppResult = StockServiceHelper.GenerateLotMasterByCodeRule(ctx, materialMetadata.BusinessInfo, lotField, array);
            var flots = codeAppResult.CodeNumbers;
            for (int i = 0; i < flots.Count; i++)
            {
                if (!flots[i].Value.IsNullOrEmptyOrWhiteSpace())
                {
                    enys[i][lotorm + "_Text"] = flots[i].Value[0];
                }
            }


        }
        /// <summary>
        /// 获取用户默认套打模板 ctx上下文，billkey单据标识，billtypeid单据类型
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="billkey"></param>
        /// <param name="fid"></param>
        /// <param name="billtypeid"></param>
        /// <returns></returns>
        public string Tdid(Context ctx, string billkey, string billtypeid)
        {
            IMetaDataService metadataService = Kingdee.BOS.App.ServiceHelper.GetService<IMetaDataService>();
            FormMetadata materialMetadata = metadataService.Load(ctx, billkey) as FormMetadata;
            string detdmb = null;
            string paratext = UserParamterServiceHelper.Load(ctx, "NotePrintSetup" + materialMetadata.BusinessInfo.GetForm().Id.ToUpper().GetHashCode().ToString(), ctx.UserId, "");//获取套打设置参数

            JArray paras = JArray.Parse(paratext);
            if (paras != null)
            {
                var tdmbs = (from pa in paras where pa["key"].ToString() == billtypeid select pa).ToList();
                if (tdmbs.Count > 0)
                {
                    detdmb = tdmbs[0]["value"].ToString();
                    return detdmb;
                }
                else
                {
                    return detdmb;
                }
            }
            else
            {
                return detdmb;
            }

        }


        /// <summary>
        ///需要引用NPOI，NPOI.OOXML，NPOI.OpenXml4Net，NPOI.OpenXmlFormats组件 datable转excel,用于查询到数据转换成excel发送给用户，data加工好的datable，sheetName excelsheet名称,isColumnWritten列名是否要写入，filename生成Excel文件的物理路径： @"E:\testfile\" + date + ".xlsx",需要定时服务器上此路径的文件
        /// </summary>
        /// <param name="data"></param>
        /// <param name="sheetName"></param>
        /// <param name="isColumnWritten"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public string DataTableToExcel(DataTable data, string sheetName, bool isColumnWritten, string fileName)
        {
            IWorkbook workbook = null;

            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;

            FileStream fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                workbook = new XSSFWorkbook();
            else if (fileName.IndexOf(".xls") > 0) // 2003版本
                workbook = new HSSFWorkbook();

            try
            {
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet(sheetName);
                }
                else
                {
                    return null;
                }

                if (isColumnWritten == true) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                    }
                    count = 1;
                }
                else
                {
                    count = 0;
                }

                for (i = 0; i < data.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }
                workbook.Write(fs); //写入到excel
                return fileName;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return null;
            }
        }
        /// <summary>
        /// 将excel文件通过邮件方式发送给用户，file Excel文件的物理路径： @"E:\testfile\" + date + ".xlsx"，sjrs收件人集合，titel邮件标题，body邮件内容,fjr发件人，fjradress发件人邮件地址，sqm授权码，host smtp服务器地址 如smtp.qq.com(qq邮箱)
        /// </summary>
        /// <param name="file"></param>
        /// <param name="sjrs"></param>
        /// <param name="title"></param>
        /// <param name="body"></param>
        /// <param name="fjr"></param>
        /// <param name="fjradress"></param>
        /// <param name="sqm"></param>
        /// <param name="host"></param>
        /// <returns></returns>
        public string SendMsg(string file, List<string> sjrs, string title, string body, string fjr, string fjradress, string sqm, string host)
        {


            System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();
            if (sjrs.Count > 0)
            {
                foreach (var sjr in sjrs)
                {
                    msg.CC.Add(sjr);//收件人
                }
                msg.From = new System.Net.Mail.MailAddress(fjr);//发件人
                msg.Subject = title;//邮件标题    
                msg.SubjectEncoding = System.Text.Encoding.UTF8;//邮件标题编码    
                msg.Body = body;//邮件内容    
                msg.BodyEncoding = System.Text.Encoding.UTF8;//邮件内容编码    
                msg.IsBodyHtml = false;//是否是HTML邮件    
                msg.Priority = MailPriority.High;//邮件优先级 
                System.Net.Mail.Attachment myAttachment = new System.Net.Mail.Attachment(file);
                System.Net.Mime.ContentDisposition disposition = myAttachment.ContentDisposition;
                disposition.CreationDate = System.IO.File.GetCreationTime(file);
                disposition.ModificationDate = System.IO.File.GetLastWriteTime(file);
                disposition.ReadDate = System.IO.File.GetLastAccessTime(file);
                msg.Attachments.Add(myAttachment);
                System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
                client.Credentials = new System.Net.NetworkCredential(fjradress, sqm);//发件人邮箱地址，授权码
                client.Host = host;
                //client.Send(msg);
                try
                {
                    client.SendAsync(msg, "");
                }
                catch (Exception ex)
                {

                    return "邮件发送失败" + ex.Message;
                }
                return "邮件发送成功！";
            }
            else
            {
                return "邮件发送失败，请添加收件人！";
            }



        }

        /// <summary>
        /// 即时库存校对方法，ctx上下文，orgs需要校对的组织，wlbms需要校对的物料编码集合（可为空）
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="orgs"></param>
        /// <param name="wlbms"></param>
        public void Jskcjd(Context ctx, List<long> orgs, List<string> wlbms)
        {
            object systemProfile = Kingdee.K3.SCM.ServiceHelper.CommonServiceHelper.GetSystemProfile(ctx, 0, "STK_StockParameter", "RecInvCheckMidData", false);
            bool recordMidData = false;
            if (systemProfile != null && !string.IsNullOrWhiteSpace(systemProfile.ToString()))
            {
                recordMidData = Convert.ToBoolean(systemProfile);
            }
            systemProfile = Kingdee.K3.SCM.ServiceHelper.CommonServiceHelper.GetSystemProfile(ctx, 0, "STK_StockParameter", "CleanReleaseLink", false);
            bool cleanReleaseLink = false;
            if (systemProfile != null && !string.IsNullOrWhiteSpace(systemProfile.ToString()))
            {
                cleanReleaseLink = Convert.ToBoolean(systemProfile);
            }

            StockServiceHelper.StockCheck(ctx, orgs, wlbms, recordMidData, cleanReleaseLink);
        }
        /// <summary>
        /// /计算表达式的方法，str表达式字符串，type动态生成的类
        /// </summary>
        /// <param name="str"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        //public decimal FieldEval(string str, Object type)
        //{
        //    //注册
        //    var reg = new TypeRegistry();

        //    reg.RegisterSymbol("Result", type);

        //    //如果要使用Math函数，还就注册这个
        //    //reg.RegisterDefaultTypes();

        //    //编译
        //    var p = new CompiledExpression(str) { TypeRegistry = reg };
        //    p.Compile();

        //    //计算
        //    Console.WriteLine("变量字段计算: {0}", p.Eval());
        //    return Convert.ToDecimal(p.Eval());
        //}

        /// 应付付款手工核销方法,ctx上下文,billnos应付单/付款单单据编号字符串,多张单据使用逗号拼接（数据库in查询格式）webapi中使用需先调用登陆接口 var ctx = this.KDContext.Session.AppContext;获取ctx,构建的ctx无效
        public string WebapiHx(Context ctx, List<string> billnos)
        {
            if (billnos.Count == 2)
            {
                var filterMetadata = FormMetaDataCache.GetCachedFilterMetaData(ctx);
                var reportFilterMetadata = FormMetaDataCache.GetCachedFormMetaData(ctx, "AP_PayMatchFilter");//
                var reportFilterServiceProvider = reportFilterMetadata.BusinessInfo.GetForm().GetFormServiceProvider();
                var model = new FilterModel();//构建过滤框业务对象
                model.SetContext(ctx, reportFilterMetadata.BusinessInfo, reportFilterServiceProvider);
                model.FormId = reportFilterMetadata.BusinessInfo.GetForm().Id;
                model.FilterObject.FilterMetaData = filterMetadata;
                List<ColumnField> ColumnFields = new List<ColumnField>();
                var filter = model.GetFilterParameter();//过滤框参数
                List<Dictionary<long, MatchBillData>> list = new List<Dictionary<long, MatchBillData>>();//核销单据集合
                Dictionary<int, string> temps = new Dictionary<int, string>();//核销临时表
                MatchParameters matchParameters = new MatchParameters();//获取核销单据参数

                for (int i = 0; i < billnos.Count; i++)
                {
                    if (i == 0)
                    {
                        filter.ColumnInfo = ColumnFields;
                        filter.FilterString = string.Format("( FSRCBILLNO in ('{0}')  )", billnos[0]);//AP00000072  FKD00000020
                        Kingdee.BOS.Orm.DataEntity.DynamicObject flter = filter.CustomFilter;
                        flter["FDateFrom"] = "1999-02-01 00:00:00";
                        flter["FBillSelect"] = "AP_Payable,AP_OtherPayable";//AP_Payable,AP_OtherPayable  AP_PAYBILL,AP_REFUNDBILL
                        flter["FSETTLEORGID"] = "1";
                        matchParameters.MatchType = "1";
                        matchParameters.BillDirection = "1";//借方单据1,贷方单据-1
                        matchParameters.SelectBillFromId = "AP_MatchBillSelect";
                        matchParameters.SpecialMatch = true;
                        matchParameters.FilterFromId = "AP_PayMatchFilter";
                        matchParameters.IsBadDebt = false;
                        Dictionary<int, string> dictionary = new Dictionary<int, string>();
                        Dictionary<int, string> temptable = new Dictionary<int, string>();
                        Dictionary<long, MatchBillData> billData = MatchServiceHelper.GetBillData(ctx, matchParameters, filter, out dictionary, temptable);
                        list.Add(billData);
                        foreach (var dic in dictionary)
                        {
                            temps.Add(dic.Key, dic.Value);
                        }
                    }
                    else
                    {
                        filter.ColumnInfo = ColumnFields;
                        filter.FilterString = string.Format("( FSRCBILLNO in ('{0}')  )", billnos[1]); ;//AP00000072  FKD00000020
                        Kingdee.BOS.Orm.DataEntity.DynamicObject flter = filter.CustomFilter;
                        flter["FDateFrom"] = "1999-02-01 00:00:00";
                        flter["FBillSelect"] = "AP_PAYBILL,AP_REFUNDBILL";//AP_Payable,AP_OtherPayable  AP_PAYBILL,AP_REFUNDBILL
                        flter["FSETTLEORGID"] = "1";
                        matchParameters.MatchType = "1";
                        matchParameters.BillDirection = "-1";//借方单据1,贷方单据-1
                        matchParameters.SelectBillFromId = "AP_MatchBillSelect";
                        matchParameters.SpecialMatch = true;
                        matchParameters.FilterFromId = "AP_PayMatchFilter";
                        matchParameters.IsBadDebt = false;
                        Dictionary<int, string> dictionary = new Dictionary<int, string>();
                        Dictionary<int, string> temptable = new Dictionary<int, string>();
                        Dictionary<long, MatchBillData> billData = MatchServiceHelper.GetBillData(ctx, matchParameters, filter, out dictionary, temptable);

                        list.Add(billData);
                        foreach (var dic in dictionary)
                        {
                            temps.Add(dic.Key, dic.Value);
                        }
                    }

                }
                MatchParameters matchParametersa = new MatchParameters();
                matchParametersa.MatchType = "1";
                matchParametersa.SpecialMatch = true;
                matchParametersa.UserId = ctx.UserId;
                matchParametersa.MatchFromId = "AP_PayMatck";
                matchParametersa.BatchGeneRecord = false;
                matchParametersa.GenBillDate = DateTime.Now;
                matchParametersa.MatchMethodID = 60;
                matchParametersa.IsBadDebt = false;

                try
                {
                    IOperationResult operationResult2 = MatchServiceHelper.SpecialMatch(ctx, matchParametersa, list, temps);
                    if (operationResult2.IsSuccess == true)//核销成功
                    {
                        return "核销成功！";
                    }
                    else
                    {
                        return operationResult2.ValidationErrors[0].Message;
                    }
                }
                catch (Exception ex)
                {

                    return ex.Message.ToString();
                }
            }
            else
            {
                return "核销单据行数有误！";
            }


        }

        /// <summary>
        /// Webapi中获取ctx方法,无需先调用登陆接口,获取到ctx后放入缓存中,使得不重复调用客户端登陆,sjzxid数据中心id,user用户,pwd密码
        /// </summary>
        /// <param name="sjzxid"></param>
        /// <param name="user"></param>
        /// <param name="pwd"></param>
        /// <returns></returns>
        public Context Logmsg(string sjzxid, string user, string pwd)
        {
            var value = HttpRuntime.Cache.Get("Context");
            if (value == null)
            {

                var proxyz = new Kingdee.BOS.ServiceFacade.KDServiceClient.User.UserServiceProxy();
                LoginInfo loginInfoz = new LoginInfo();
                loginInfoz.AcctID = sjzxid;
                loginInfoz.Username = user;
                loginInfoz.Password = pwd;
                loginInfoz.Lcid = 2052;
                var ctx = proxyz.ValidateUser("", loginInfoz).Context;
                try
                {
                    object obj1 = HttpRuntime.Cache.Add("Context", ctx, null, DateTime.Now.AddMinutes(25), TimeSpan.Zero, CacheItemPriority.Default, null);//将上下文信息添加到缓存中并设置到期时间25分钟
                    var value2 = HttpRuntime.Cache.Get("Context");//获取缓存中的当前上下文
                    return (Context)value2;
                }
                catch (Exception)
                {

                    return null;
                }


            }
            else
            {
                return (Context)value;
            }
        }

        /// <summary>
        /// 触发工作流方法，ctx上下文，businessInfo 元数据，lcid流程模板id，lcmxid流程模板明细id,lcbb流程版本,fid单据内码，fqr发起人
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="businessInfo"></param>
        /// <param name="lcid"></param>
        /// <param name="lcmxid"></param>
        /// <param name="lcbb"></param>
        /// <param name="fid"></param>
        /// <param name="fqr"></param>
        /// <returns></returns>
        public bool TiggerWorkFlow(Context ctx, BusinessInfo businessInfo, string lcid, string lcmxid, string lcbb, string fid, int fqr)
        {
            List<Kingdee.BOS.Orm.DataEntity.DynamicObject> lists = new List<Kingdee.BOS.Orm.DataEntity.DynamicObject>();
            var sl = this.CreateProcessInstance(ctx, businessInfo, lcid, lcmxid, lcbb, fid, false, lists, fqr, null, "");

            bool result = this.StartWorkflow(ctx, false, sl, this.GetMessageSendMode(ctx));
            return result;
        }

        private ProcessInstance CreateProcessInstance(Context ctx, BusinessInfo info, string templeteId, string tempDetailId, string verId, object pkValue, bool writeProcessLog, List<Kingdee.BOS.Orm.DataEntity.DynamicObject> storeList, int? originatorPostId = null, string submitOpinion = null, string submitStatusKey = "")
        {

            WorkflowModelService workflowModelService = new WorkflowModelService();
            CacheProcess cacheProcessByVersionId = workflowModelService.GetCacheProcessByVersionId(ctx, verId);
            if (cacheProcessByVersionId == null)
            {
                throw new Exception(ResManager.LoadKDString("未找到对应的流程定义，版本Id：", "002525030016300", SubSystemType.BOS, new object[0]) + verId);
            }
            if (info == null)
            {
                throw new Exception(ResManager.LoadKDString("_businessInfo 未赋值。", "002525030021211", SubSystemType.BOS, new object[0]));
            }
            BOSFlowRepository bosflowRepository = new BOSFlowRepository(ctx);
            WorkflowEngine workflowEngine = new WorkflowEngine(cacheProcessByVersionId.Process, ctx, bosflowRepository);
            ProcessInstance processInstance = workflowEngine.CreateProcessInstance();
            processInstance.TrySetMember("FormId", info.GetForm().Id);
            processInstance.TrySetMember("FormName", info.GetForm().Name.ToString());
            processInstance.TrySetMember("KeyValue", pkValue.ToString());
            processInstance.ProcDefId = cacheProcessByVersionId.ProcDefId;
            processInstance.VersionId = verId;
            processInstance.OriginatorId = ctx.UserId;
            processInstance.TempleteId = templeteId;
            processInstance.TempDetailId = tempDetailId;
            ProcessInstance processInstance2 = processInstance;
            int? num = originatorPostId;
            processInstance2.OriginatorPostId = ((num != null) ? new long?((long)num.GetValueOrDefault()) : null);
            processInstance.CreatedTime = DateTime.Now;
            processInstance.Summary = "123";
            processInstance.SubmitOpinion = submitOpinion;
            processInstance.SubmitStatusKey = ((submitStatusKey == null) ? "" : submitStatusKey);
            this.InitProcInst(processInstance, ctx);
            bosflowRepository.CreateProcessInstance(processInstance);
            string workflowStartActivityCheckId = this.GetWorkflowStartActivityCheckId(info.GetForm().Id, pkValue);
            this.AddStartBackMqStore(ctx, processInstance, workflowStartActivityCheckId, storeList);
            return processInstance;
        }

        public string GetWorkflowStartActivityCheckId(string formId, object pkValue)
        {
            return formId + "_" + pkValue.ToString();
        }

        public void InitProcInst(ProcessInstance procInst, Context ctx)
        {
            object obj;
            procInst.TryGetMember("FormId", out obj);
            object obj2;
            procInst.TryGetMember("KeyValue", out obj2);
            if (obj == null || string.IsNullOrWhiteSpace(obj.ToString()) || obj2 == null || string.IsNullOrWhiteSpace(obj2.ToString()))
            {
                throw new Exception(ResManager.LoadKDString("流程实例的单据变量和单据编码变量不能为空", "002525030016507", SubSystemType.BOS, new object[0]));
            }
            procInst.Number = "CGSSD" + "_" + DateTime.Now.ToString("yyyy_MM_dd");//实例名称，一定要包含"_"
            string strSQL = "\r\n                insert into T_WF_PIBIMap(FId, FProcInstId, FObjectTypeId, FKeyValue)\r\n                values(@FId, @FProcInstId, @FObjectTypeId, @FKeyValue)";
            IDBService service = ServiceHelper.GetService<IDBService>();
            List<SqlParam> list = new List<SqlParam>();
            list.Add(new SqlParam("@FId", KDDbType.AnsiString, service.GetSequenceString(1)[0]));
            list.Add(new SqlParam("@FProcInstId", KDDbType.AnsiString, procInst.MapInstanceId));
            list.Add(new SqlParam("@FObjectTypeId", KDDbType.AnsiString, obj.ToString()));
            list.Add(new SqlParam("@FKeyValue", KDDbType.AnsiString, obj2.ToString()));
            DBUtils.Execute(ctx, strSQL, list);
        }

        public bool StartWorkflow(Context ctx, bool writeProcessLog, ProcessInstance pi, MessageSendMode messageSendMode)
        {


            if (messageSendMode == MessageSendMode.Sync || messageSendMode == MessageSendMode.Async)
            {

                try
                {
                    this.SendSyncMessage<StartBackJobService>(ctx, pi, null, true);
                }
                catch (Exception)
                {

                    return false;
                }
                return true;
            }
            else
            {
                try
                {
                    this.SendAsyncMessage<StartBackJobService>(ctx, messageSendMode, pi, null);
                }
                catch (Exception)
                {

                    return false;
                }
                return true;

            }

        }
        public void AddStartBackMqStore(Context ctx, ProcessInstance pi, string startActivityInstId, List<Kingdee.BOS.Orm.DataEntity.DynamicObject> storeList)
        {
            AsyncMessage msg = new AsyncMessage
            {
                Context = ctx,
                ServiceType = typeof(StartBackJobService).AssemblyQualifiedName,
                DataInfo = pi
            };
            this.AddMqStoreList(ctx, pi.CreatedTime, storeList, startActivityInstId, msg);
        }
        public void AddMqStoreList(Context ctx, DateTime executedTime, List<Kingdee.BOS.Orm.DataEntity.DynamicObject> storeList, string actInstId, AsyncMessage msg)
        {
            storeList.Add(new MapStateMqStore
            {
                MsgData = SerializatonUtil.Serialize(msg),
                SendUserId = ctx.UserId,
                CREATEDATE = executedTime,
                ActInstId = actInstId,
                Status = MapStateMqStatus.Untreated
            }.DataEntity);
        }
        public Exception SendAsyncMessage<T>(Context ctx, MessageSendMode mode, object dataInfo, Action method)
        {
            AsyncMessage msg = new AsyncMessage
            {
                Context = ctx,
                DataInfo = dataInfo,
                ServiceType = typeof(T).AssemblyQualifiedName
            };
            return this.SendAsyncMessage(msg, mode, method);
        }
        public Exception SendAsyncMessage(AsyncMessage msg, MessageSendMode mode, Action method)
        {
            Exception result = null;
            if (method != null)
            {
                method();
            }
            try
            {
                MessageQueueAsyncService messageQueueAsyncService = new MessageQueueAsyncService();
                messageQueueAsyncService.Send(msg);
            }
            catch (Exception ex)
            {
                result = ex;
                Logger.Error("WF", ex.Message, ex);
            }
            return result;
        }

        public void SendSyncMessage<T>(Context ctx, object dataInfo, Action method, bool startNewThread = true)
        {
            if (startNewThread)
            {
                MainWorker.QuequeTask(ctx, delegate ()
                {
                    T t2 = Activator.CreateInstance<T>();
                    if (method != null)
                    {
                        method();
                    }
                    ((IAsyncService)((object)t2)).Excute(ctx, dataInfo);
                }, delegate (AsynResult result)
                {
                    if (!result.Success && result.Exception != null)
                    {

                    }
                });
                return;
            }
            try
            {
                T t = Activator.CreateInstance<T>();
                using (KDTransactionScope kdtransactionScope = new KDTransactionScope(TransactionScopeOption.Required))
                {
                    if (method != null)
                    {
                        method();
                    }
                    ((IAsyncService)((object)t)).Excute(ctx, dataInfo);
                    kdtransactionScope.Complete();
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public MessageSendMode GetMessageSendMode(Context ctx)
        {
            bool sysParaWriteProcessLogger = this.GetSysParaWriteProcessLogger(ctx);
            //   LogRepository.WriteBOSWorkflowLog(ctx, "GetMessageSendMode", 10, "begin GetMessageSendMode", sysParaWriteProcessLogger);
            SystemParameterService systemParameterService = new SystemParameterService();
            object paramter = systemParameterService.GetParamter(ctx, 0L, 0L, "WF_SystemParameter", "FlowProcMode", 0L);
            //    LogRepository.WriteBOSWorkflowLog(ctx, "GetMessageSendMode", 20, string.Format("FlowProcMode = {0}", paramter), sysParaWriteProcessLogger);
            MessageSendMode messageSendMode = (MessageSendMode)Convert.ToInt32(paramter);
            if (messageSendMode != MessageSendMode.MSMQ)
            {
                return messageSendMode;
            }
            else
            {
                return MessageSendMode.Async;
            }
        }

        public bool GetSysParaWriteProcessLogger(Context ctx)
        {
            return ObjectUtils.Object2Bool(ServiceHelper.GetService<ISystemParameterService>().GetParamter(ctx, 0L, 0L, "WF_SystemParameter", "FWriteLog", 0L), false);
        }
        public class StartBackJobService : IAsyncService
        {

            public void Excute(Context ctx, object arg)
            {

                this.StartWorkflowExecute(ctx, (ProcessInstance)arg);
            }


            public string GetDataInfo(object arg)
            {
                ProcessInstance processInstance = (ProcessInstance)arg;
                StringBuilder stringBuilder = new StringBuilder();
                stringBuilder.AppendFormat(ResManager.LoadKDString("流程实例Id:{0}, 流程编码:{1}, 发起人Id:{2}", "002525030023196", SubSystemType.BOS, new object[0]), processInstance.MapInstanceId, processInstance.Number, processInstance.OriginatorId.ToString());
                stringBuilder.Append(" \r\n");
                foreach (VariableInstance variableInstance in processInstance.VariableInstances)
                {
                    stringBuilder.AppendFormat(ResManager.LoadKDString("变量名:{0}, 值:{1}", "002525030023199", SubSystemType.BOS, new object[0]), variableInstance.Name, variableInstance.Value);
                    stringBuilder.Append(" \r\n");
                }
                return stringBuilder.ToString();
            }

            public void StartWorkflowExecute(Context ctx, ProcessInstance procInst)
            {

                WorkflowModelService workflowModelService = new WorkflowModelService();
                CacheProcess cacheProcessByVersionId = workflowModelService.GetCacheProcessByVersionId(ctx, procInst.VersionId);
                BOSFlowRepository repository = new BOSFlowRepository(ctx);
                object obj = "";
                procInst.TryGetMember("FormId", out obj);
                object obj2 = "";
                procInst.TryGetMember("KeyValue", out obj2);
                using (KDTransactionScope kdtransactionScope = new KDTransactionScope(TransactionScopeOption.Required))
                {

                    WorkflowEngine workflowEngine = new WorkflowEngine(cacheProcessByVersionId.Process, ctx, repository);
                    IDictionary<string, object> variables = null;
                    workflowEngine.Start(procInst, variables);
                    kdtransactionScope.Complete();
                }
            }
        }
    }
    /// <summary>
    /// 动态生成类，用于动态嵌入表达式
    /// </summary>
    //public class Dtfileds : System.Dynamic.DynamicObject
    //{
    //    public string PropertyName { get; set; }
    //    Dictionary<string, object> Properties = new Dictionary<string, object>();
    //    public override bool TrySetMember(SetMemberBinder binder, object value)
    //    {
    //        if (binder.Name == "Property")
    //        {
    //            Properties[PropertyName] = value;
    //        }
    //        else
    //        {
    //            Properties[binder.Name] = value;
    //        }
    //        return true;
    //    }
    //    public override bool TryGetMember(GetMemberBinder binder, out object result)
    //    {
    //        var keys = (from Propertie in Properties where Propertie.Key.Trim() == binder.Name select Propertie).ToList();
    //        if (keys.Count > 0)
    //        {
    //            return Properties.TryGetValue(keys[0].Key, out result);
    //        }

    //        result = null;
    //        return false;
    //    }
    //}


}
