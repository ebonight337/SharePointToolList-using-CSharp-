using System;
using System.Net;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Threading;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;

namespace MaintenanceListData
{
    /// <summary>
    /// SharePoint のデータ操作機能を提供します
    /// </summary>
    public class SharePointDataService : IDisposable
    {
        private readonly SP.ClientContext _context;
        private readonly SecureString _secureString;

        /// <summary>
        /// 操作対象の ClientContext を取得します
        /// </summary>
        public SP.ClientContext ClientContext => _context;
        /// <summary>
        /// ジョブの詳細ログを取得します
        /// </summary>
        static Logger jobLog = new Logger("Log//SP_DetailLog_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt", System.Text.Encoding.UTF8);

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="sitePath">データ操作対象のサイト URL</param>
        /// <param name="mailAddress">データ操作を行うユーザーメールアドレス</param>
        /// <param name="password">データ操作を行うユーザーパスワード</param>
        public SharePointDataService(string sitePath, string mailAddress, string password)
        {
            _context = new SP.ClientContext(sitePath);

            _secureString = new SecureString();
            password.ToList().ForEach(c => _secureString.AppendChar(c));
            _context.Credentials = new SP.SharePointOnlineCredentials(mailAddress, _secureString);
        }




        /// <summary>
        /// 対象の名称のリストを取得します
        /// </summary>
        /// <param name="listName"></param>
        /// <returns></returns>
        public SP.List GetListByName(string listName)
        {
            jobLog.Write("リスト名を取得します => " + listName);
            return _context.Web.Lists.GetByTitle(listName);
        }

        /// <summary>
        /// ログイン名（メールアドレス）から有効なユーザー情報を取得します
        /// </summary>
        /// <param name="email"></param>
        /// <returns></returns>
        public SP.User GetUser(string email)
        {
            jobLog.Write("ユーザープリンシパル名を取得します => " + email);
            return _context.Web.EnsureUser(email);
        }

        public SP.UserCollection GetAllUser()
        {
            jobLog.Write("全ユーザーを取得します");
            SP.UserCollection coll = _context.Web.SiteUsers;
            _context.Load(coll);
            _context.ExecuteQuery();
            return coll;
        }

        /// <summary>
        /// 対象のリストの URL を取得します
        /// </summary>
        /// <param name="spList"></param>
        /// <returns></returns>
        public async Task<string> GetListUrl(SP.List spList)
        {
            if (spList == null)
            {
                throw new ArgumentNullException(nameof(spList));
            }

            jobLog.Write("リストURLを取得します");
            _context.Load(spList, l => l.RootFolder.ServerRelativeUrl);
            try
            {
                await _context.ExecuteQueryAsync();
                jobLog.Write(spList.RootFolder.ServerRelativeUrl);
                return spList.RootFolder.ServerRelativeUrl;
            }
            catch
            {
                return null;
            }
        }

        public async Task<string> GetViewQuery(SP.List spList)
        {
            if (spList == null)
            {
                throw new ArgumentNullException(nameof(spList));
            }

            var view = spList.DefaultView;

            if (view == null)
            {
                throw new ArgumentNullException(nameof(spList));
            }

            _context.Load(view, v => v.ViewQuery);
            await _context.ExecuteQueryAsync();

            return view.ViewQuery;
        }

        public async Task<SP.ListItemCollection> GetListItemTop10(SP.List spList)
        {
            SP.CamlQuery query = new SP.CamlQuery();
            query.ViewXml = "<View><Query></Query><ViewFields><FieldRef Name='Title' /></ViewFields><RowLimit>10</RowLimit></View>";
            SP.ListItemCollection collListItem = spList.GetItems(query);

            _context.Load(collListItem);
            await _context.ExecuteQueryAsync();

            return collListItem;
        }

        public async Task<SP.ListItem> GetListItemByTitle(SP.List spList, string itemTitle, string folderPath)
        {
            SP.CamlQuery query = new SP.CamlQuery();
            query.ViewXml = $"<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>{itemTitle}</Value></Eq></Where></Query></View>";
            if (!string.IsNullOrEmpty(folderPath)) query.FolderServerRelativeUrl = folderPath;
            SP.ListItemCollection collListItem = spList.GetItems(query);

            _context.Load(collListItem);
            await _context.ExecuteQueryAsync();

            return collListItem.Count > 0 ? collListItem[0] : null;
        }


        /// <summary>
        /// リストを作成します
        /// </summary>
        /// <param name="listName"></param>
        /// <param name="Description"></param>
        /// <param name="targetListTemplate"></param>
        /// <returns></returns>
        public async Task<SP.List> CreateListAsync(string listName, string targetListTemplate, string listUrl, string Description=null)
        {
            jobLog.Write("リストの作成を実施します => " + listName);
            try
            {
                // 存在確認
                SP.ListCollection lists = _context.Web.Lists;
                _context.Load(lists);
                await _context.ExecuteQueryAsync();

                foreach(SP.List oList in lists)
                {
                    if (oList.Title == listName)
                    {
                        jobLog.Write("作成しようとしているリストはすでに存在します =>" + listName);
                        throw new Exception("作成しようとしているリストはすでに存在します =>" + listName);
                    }
                }

                // リストテンプレートを取得
                SP.ListTemplateCollection templates = _context.Site.GetCustomListTemplates(_context.Web);
                _context.Load(templates);
                await _context.ExecuteQueryAsync();

                SP.ListCreationInformation listCreationInfo = new SP.ListCreationInformation();
                listCreationInfo.Title = listUrl;
                if(Description != null)
                {
                    listCreationInfo.Description = Description;
                }
                
                SP.ListTemplate listTemplate = templates.First(listTemp => listTemp.Name == targetListTemplate);
                listCreationInfo.ListTemplate = listTemplate;
                listCreationInfo.TemplateFeatureId = listTemplate.FeatureId;
                listCreationInfo.TemplateType = listTemplate.ListTemplateTypeKind;
                var createList = _context.Web.Lists.Add(listCreationInfo);

                // リスト名変更
                createList.Title = listName;
                createList.Update();
                await _context.ExecuteQueryAsync();

            }
            catch(Exception e)
            {
                jobLog.Write("リストの作成中にエラーが発生しましたエラーを確認してください\r\n" + e);
                throw new Exception("リストの作成でエラーが発生しましたエラーを確認してください");
            }
            try
            {
                jobLog.Write("リストの作成が完了しました => " + listName);
                return _context.Web.Lists.GetByTitle(listName);
            }
            catch(Exception e)
            {
                jobLog.Write("作成したリストを取得できませんでしたエラーを確認してください\r\n" + e);
                throw new Exception("作成したリストを取得できませんでしたエラーを確認してください");
            }
        }

        public async Task<bool> DeleteListAsync(string listName)
        {
            jobLog.Write("リストの削除を実施します => " + listName);
            // 存在確認
            SP.ListCollection lists = _context.Web.Lists;
            try
            {
                _context.Load(lists);
                await _context.ExecuteQueryAsync();
            }catch(Exception e)
            {
                jobLog.Write("リストの取得でエラーが発生しました。\r\n" + e);
                throw new Exception("リストの取得でエラーが発生しました。エラーを確認してください。");
            }

            foreach (SP.List oList in lists)
            {
                if (oList.Title == listName)
                {
                    SP.List deleteList = _context.Web.Lists.GetByTitle(listName);
                    try
                    {
                        deleteList.DeleteObject();
                        await _context.ExecuteQueryAsync();
                    }
                    catch (Exception e)
                    {
                        jobLog.Write("リスト削除処理で以下のエラーが発生しました\r\n" + e);
                        // throw new Exception("リスト削除処理で以下のエラーが発生しましたログを確認してください");
                    }
                    jobLog.Write("リストの削除が完了しました => " + listName);
                    return true;
                }
            }

            jobLog.Write("削除対象のリストが存在しないため、削除処理をパスします =>" + listName);
            Console.WriteLine("削除対象のリストが存在しないため、削除処理をパスします =>" + listName);
            return true;

        }

        /// <summary>
        /// 対象のリストにアイテムを非同期処理で追加します
        /// </summary>
        /// <param name="spList">追加対象のSharePointList</param>
        /// <param name="item">追加するアイテム</param>
        /// <param name="attachmentFilePaths">追加する添付ファイルパスリスト</param>
        /// <param name="folderPath">追加する対象フォルダ（省略した場合はリストのルートフォルダ）</param>
        /// <returns></returns>
        public async Task<object> AddItemToListAsync(SP.List spList, IDictionary<string, object> item, IEnumerable<string> attachmentFilePaths, string folderPath = null)
        {
            jobLog.Write($"アイテム作成開始=> {item["Title"]}");
            if (spList == null)
            {
                throw new ArgumentNullException(nameof(spList));
            }
            if (item == null)
            {
                throw new ArgumentNullException(nameof(item));
            }

            // フォルダの指定
            var listItemCreationInfo = new SP.ListItemCreationInformation();
            if (!string.IsNullOrEmpty(folderPath))
            {
                listItemCreationInfo.FolderUrl = folderPath;
            }

            var newItem = spList.AddItem(listItemCreationInfo);


            // FileStream リソース管理用リスト
            //  → ExecuteQueryAsync 呼び出し前に FileStream を using で解放してしまうと、
            //      ExecuteQueryAsync の時点でファイル参照できなくなってしまうため、明示的にリソース管理
            var fileStreamList = new List<FileStream>();

            newItem["_ModerationStatus"] = 0;
            newItem.UpdateOverwriteVersion();

            try
            {
                if (attachmentFilePaths != null)
                {
                    // 添付ファイルの追加処理
                    foreach (var filePath in attachmentFilePaths)
                    {
                        if (!File.Exists(filePath))
                        {
                            newItem.DeleteObject();
                            throw new NullReferenceException($"指定されたファイルが存在しません => {filePath}");
                        }

                        var fileStream = new FileStream(filePath, FileMode.Open);

                        if (fileStream.Length == 0)
                        {
                            newItem.DeleteObject();
                            throw new NullReferenceException($"ファイル容量が0です => {filePath}");
                        }

                        fileStreamList.Add(fileStream);
                        var attachmentInfo = new SP.AttachmentCreationInformation
                        {
                            ContentStream = fileStream,
                            FileName = Path.GetFileName(filePath)
                        };

                        newItem.AttachmentFiles.Add(attachmentInfo);
                    }

                }

                // アイテム追加処理(添付ファイル前に実施すると、更新者が実行アカウントになってしまう)
                foreach (var keyValue in item)
                {
                    // 作成日時がない場合は現在の日時挿入
                    if ((keyValue.Key.ToString() == "Created") && ((keyValue.Value.Equals(""))))
                    {
                        newItem[keyValue.Key] = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    }
                    else
                    {
                        newItem[keyValue.Key] = keyValue.Value;
                    }
                    // ColumNameとしてCreatedを入れているのにnullにすることはできない
                    // if ((keyValue.Key.ToString() == "Created") && ((keyValue.Value.Equals(""))))
                    // {
                    //     newItem.DeleteObject();
                    //     throw new Exception("アイテムの作成日時プロパティが不足しています");
                    // }
                }

                newItem.Update();
                // サーバーにリクエスト(429、503のエラーが発生するとリトライを行う)
                await ExecuteQuery.ExecuteQueryWithIncrementalRetry(_context);
                jobLog.Write($"アイテム作成完了=> {item["Title"]}");
            }
            catch (Exception e)
            {
                Console.WriteLine($"アイテム作成中にエラーが発生しました 詳細ログを確認してください");
                jobLog.Write("アイテム作成中に以下のエラーが発生しました\r\n" + e);
                return null;
            }
            finally
            {
                // 添付ファイルの後処理
                foreach (var fileStream in fileStreamList)
                {
                    fileStream.Dispose();
                }
            }

            return newItem;
        }

        /// <summary>
        /// 対象のリストにアイテムを非同期処理で追加します
        /// </summary>
        /// <param name="spListItem">更新対象のSharePointListItem</param>
        /// <param name="item">更新するアイテム</param>
        /// <param name="attachmentFilePaths">追加する添付ファイルパスリスト</param>
        /// <returns></returns>
        public async Task<SP.ListItem> UpdateItemToListAsync(SP.ListItem spListItem, IDictionary<string, object> item, IEnumerable<string> attachmentFilePaths)
        {
            // FileStream リソース管理用リスト
            //  → ExecuteQueryAsync 呼び出し前に FileStream を using で解放してしまうと、
            //      ExecuteQueryAsync の時点でファイル参照できなくなってしまうため、明示的にリソース管理
            var fileStreamList = new List<FileStream>();

            try
            {
                    if (spListItem == null)
                {
                    throw new ArgumentNullException(nameof(spListItem));
                }
                if (item == null)
                {
                    throw new ArgumentNullException(nameof(item));
                }

                foreach (var keyValue in item)
                {
                    spListItem[keyValue.Key] = keyValue.Value;
                }
                spListItem.Update();

            
                if (attachmentFilePaths != null)
                {



                    // 添付ファイルの更新処理
                    foreach (var filePath in attachmentFilePaths)
                    {
                        if (!File.Exists(filePath))
                        {
                            throw new InvalidOperationException($"指定されたファイルが存在しませんFilePath : {filePath}");
                        }

                        var fileTitle = Path.GetFileName(filePath);


                        // 添付ファイルの削除
                        //SP.AttachmentCollection AttachmentCollection = spListItem.AttachmentFiles;
                        //SP.Attachment existAttachment = AttachmentCollection.GetByFileName(fileTitle);
                        //_context.Load(AttachmentCollection);
                        //_context.Load(existAttachment, Attachmenttitle => Attachmenttitle);
                        //await _context.ExecuteQueryAsync();
                        //Console.WriteLine("AttachmentFiles count is: {0}", AttachmentCollection.Count);

                        //if (AttachmentCollection.Count != 0)
                        //{

                        //    existAttachment.DeleteObject();

                        //}
                        //else
                        //{
                        //    jobLog.Write("添付ファイルがないため、パスします");
                        //    Console.WriteLine("添付ファイルがないため、パスします");
                        //}



                        // 添付ファイルの追加
                        var fileStream = new FileStream(filePath, FileMode.Open);
                        fileStreamList.Add(fileStream);

                        var attachmentInfo = new SP.AttachmentCreationInformation
                        {
                            ContentStream = fileStream,
                            FileName = fileTitle
                        };
                        spListItem.AttachmentFiles.Add(attachmentInfo);
                    }

                }

                // サーバーにリクエスト(429、503のエラーが発生するとリトライを行う)
                await ExecuteQuery.ExecuteQueryWithIncrementalRetry(_context);
            }
            catch(Exception e)
            {
                Console.WriteLine($"アイテム更新中にエラーが発生しました 詳細ログを確認してください");
                jobLog.Write("アイテム更新中に以下のエラーが発生しました\r\n" + e);
                return null;
            }
            finally
            {
                // 添付ファイルの後処理
                foreach (var fileStream in fileStreamList)
                {
                    fileStream.Dispose();
                }
            }

            return spListItem;
        }

        /// <summary>
        /// 対象のリストにフォルダを非同期処理で追加します
        /// </summary>
        /// <param name="spList">リスト</param>
        /// <param name="folderPath">追加するフォルダパス</param>
        /// <param name="folderName">フォルダ名</param>
        /// <param name="createUser">フォルダ作成者</param>
        /// <returns>追加したフォルダアイテム</returns>
        public async Task<object> AddFolderToListAsync(SP.List spList, string folderPath, string folderName, SP.User createUser = null)
        {
            if (spList == null)
            {
                throw new ArgumentNullException(nameof(spList));
            }

            var folderUrls = folderName.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            jobLog.Write("フォルダを作成します　=>  " + folderUrls[0]);

            SP.ListItemCreationInformation newItem = new SP.ListItemCreationInformation();
            newItem.UnderlyingObjectType = SP.FileSystemObjectType.Folder;
            newItem.LeafName = folderUrls[0];
            newItem.FolderUrl = folderPath;
            SP.ListItem item = spList.AddItem(newItem);
            item["Title"] = folderUrls[0];

            if (createUser != null)
            {
                item["Author"] = createUser;
            }

            try
            {
                item.Update();
                // なぜかExecuteQuery関数が利用できないのでそのまま実行
                await _context.ExecuteQueryAsync();
            }
            catch (Microsoft.SharePoint.Client.ServerException e)
            {
                if (e.ServerErrorTypeName == "Microsoft.SharePoint.SPException")
                {
                    jobLog.Write("フォルダが既に存在するため、スキップします => " + folderUrls[0]);
                }
            }
            catch(Exception e)
            {
                jobLog.Write("フォルダ作成中に例外が発生しました。ログを確認してください。\r\n" + e);
                throw new Exception("フォルダ作成中に例外が発生しました。ログを確認してください。");
            }

            if (folderUrls.Length > 1)
            {
                folderPath = folderPath + "/" + folderUrls[0];
                var subFolderName = string.Join("/", folderUrls, 1, folderUrls.Length - 1);
                return await AddFolderToListAsync(spList, folderPath, subFolderName, createUser);
            }
            return item;
        }

        /// <summary>
        /// 対象のアイテムに権限を非同期処理で付与します
        /// </summary>
        /// <param name="item">付与対象のアイテム</param>
        /// <param name="groupNames">付与するグループ</param>
        /// <param name="roleTypes">付与するロール</param>
        /// <returns></returns>
        public async Task AddRoleAsync(SP.SecurableObject item, IEnumerable<string> groupNames, IEnumerable<SP.RoleType> roleTypes)
        {
            if (item == null)
            {
                throw new ArgumentNullException(nameof(item));
            }
            if (groupNames == null)
            {
                throw new ArgumentNullException(nameof(groupNames));
            }
            if (roleTypes == null)
            {
                throw new ArgumentNullException(nameof(roleTypes));
            }

            var roles = new SP.RoleDefinitionBindingCollection(_context);
            foreach (var roleType in roleTypes)
            {
                roles.Add(_context.Web.RoleDefinitions.GetByType(roleType));
            }

            item.BreakRoleInheritance(false, false);
            foreach (var groupName in groupNames)
            {
                item.RoleAssignments.Add(_context.Web.SiteGroups.GetByName(groupName), roles);
            }

            await _context.ExecuteQueryAsync();
        }
        /// <summary>
        /// リストのリネーム
        /// </summary>
        /// <param name="spList">リネーム対象のリスト</param>
        /// <param name="newName">新しいリストのタイトル</param>
        /// <param name="newUrl">新しいリストのURL</param>
        /// <param name="sitePath">サイトのパス</param>
        /// <returns></returns>
        public async Task RenameListUrl(string spList, string newName, string newUrl, string sitePath)
        {
            jobLog.Write("リネーム処理を開始します => " + spList);

            if (spList != null && newUrl != null)
            {
                SP.ListCollection lists = _context.Web.Lists;
                lists.RefreshLoad();
                await _context.ExecuteQueryAsync();

                SP.List targetList = lists.GetByTitle(spList);
                targetList.Update();
                await _context.ExecuteQueryAsync();
                try
                {
                    // /sites/testenv/Lists/temp1/Lists/backuplist
                    jobLog.Write("対象リストURL => " + await GetListUrl(targetList));
                    jobLog.Write("移行後のURL => " + sitePath + "/Lists/" + newUrl);
                    await _context.ExecuteQueryAsync();

                    targetList.RootFolder.MoveTo(sitePath + "/Lists/" + newUrl);
                    targetList.Update();
                    await _context.ExecuteQueryAsync();

                    targetList.Title = newName;
                    targetList.Update();
                    await _context.ExecuteQueryAsync();

                }catch(Exception e)
                {
                    jobLog.Write("リネーム処理でエラーが発生しましたログを確認してください\r\n" + e);
                    throw new ArgumentNullException("リネーム処理でエラーが発生しましたログを確認してください");
                }
            }
            else
            {
                jobLog.Write("リネーム処理でエラーが発生しました引数を確認してください");
                throw new ArgumentNullException("リネーム処理でエラーが発生しました引数を確認してください");
            }

            jobLog.Write("リストのリネームが完了しました => " + newUrl);
        }

        public async Task<bool> NoCrawlSetting(string spList)
        {
            var target = GetListByName(spList);
            //Console.WriteLine(await GetListUrl(target));
            try
            {
                // nocrawl 処理
                target.NoCrawl = true;
                target.Update();
                await _context.ExecuteQueryAsync();
            }
            catch (Exception e)
            {
                jobLog.Write("検索結果に反映しない設定でエラーが発生しました \r\n" + e);
                // throw new Exception("検索結果に反映しない設定でエラーが発生しました \r\n" + e);
            }

            return true;
        }

        /// <summary>
        /// インスタンスの破棄処理を行います
        /// </summary>
        public void Dispose()
        {
            _context?.Dispose();
            _secureString?.Dispose();
        }



    } // End _context

    static class ExecuteQuery
    {
        /// <summary>
        /// ExecuteQueryを実行する際に挟むメソッドSPの「調整」対策
        /// </summary>
        /// <param name="clientContext"></param>
        /// <param name="retryCount"></param>
        /// <param name="delay"></param>
        public static async Task ExecuteQueryWithIncrementalRetry(this SP.ClientContext clientContext, int retryCount = 5, int delay = 30000)
        {
            int retryAttempts = 0;
            int backoffInterval = delay;
            int retryAfterInterval = 0;
            bool retry = false;
            SP.ClientRequestWrapper wrapper = null;
            if (retryCount <= 0)
                throw new ArgumentException("0より大きい試行回数を設定してください");
            if (delay <= 0)
                throw new ArgumentException("0より大きい遅延を設定してください");

            // 再試行の再試行回数をモニター
            while (retryAttempts < retryCount)
            {
                try
                {
                    if (!retry)
                    {
                        await clientContext.ExecuteQueryAsync();
                        return;
                    }
                    else
                    {
                        // リクエストを再試行する
                        if (wrapper != null && wrapper.Value != null)
                        {
                            clientContext.RetryQuery(wrapper.Value);
                            return;
                        }
                    }
                }
                catch (WebException ex)
                {
                    var response = ex.Response as HttpWebResponse;
                    // Check if request was throttled - http status code 429
                    // Check is request failed due to server unavailable - http status code 503
                    if (response != null && (response.StatusCode == (HttpStatusCode)429 || response.StatusCode == (HttpStatusCode)503))
                    {
                        wrapper = (SP.ClientRequestWrapper)ex.Data["ClientRequest"];
                        retry = true;

                        // 利用可能な場合には再試行後のヘッダーを使用する
                        string retryAfterHeader = response.GetResponseHeader("Retry-After");
                        if (!string.IsNullOrEmpty(retryAfterHeader))
                        {
                            if (!Int32.TryParse(retryAfterHeader, out retryAfterInterval))
                            {
                                retryAfterInterval = backoffInterval;
                            }
                        }
                        else
                        {
                            retryAfterInterval = backoffInterval;
                        }

                        // 指定されたミリ秒分遅らせる
                        Thread.Sleep(retryAfterInterval);

                        // Increase counters
                        retryAttempts++;
                        backoffInterval = backoffInterval * 2;
                    }
                    else
                    {
                        throw;
                    }
                }
            }
            throw new MaximumRetryAttemptedException($"Maximum retry attempts {retryCount}, has be attempted.");
        }

        [Serializable]
        public class MaximumRetryAttemptedException : Exception
        {
            public MaximumRetryAttemptedException(string message) : base(message) { }
        }

    }

}
