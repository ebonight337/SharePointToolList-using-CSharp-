using System;
using System.Net;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Threading;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;

namespace SPOSPGroupMaintenanceTool
{
    /// <summary>
    /// SharePoint のデータ操作機能を提供します
    /// </summary>
    public class SharePointDataService : IDisposable
    {
        private readonly SP.ClientContext _context;
        private readonly SecureString _secureString;
        static Logger jobLog = new Logger("Log//AuthSync_DetailLog_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt", System.Text.Encoding.UTF8);

        public SP.ClientContext ClientContext => _context;


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
        /// SPグループ状況確認
        /// </summary>
        /// <param name="weburl"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        public async Task<List<string>> GetSPWebRoleAssignments(string weburl, string username, string password)
        {
            try
            {

                jobLog.Write("=== サイトに権限が付与されている全ユーザーを取得します。 => " + weburl +" ===");
                SP.Web RootWeb = _context.Web;
                _context.Load(RootWeb, w => w.HasUniqueRoleAssignments, w => w.Url, w => w.ServerRelativeUrl);
                _context.Load(RootWeb.RoleAssignments);
                await _context.ExecuteQueryWithIncrementalRetry();

                List<string> sb = new List<string>();
                foreach (SP.RoleAssignment ra in RootWeb.RoleAssignments)
                {
                    _context.Load(ra.Member);
                    await _context.ExecuteQueryWithIncrementalRetry();

                    jobLog.Write(ra.Member.LoginName + " : " + ra.Member.PrincipalType);

                    if (ra.Member.PrincipalType.ToString() == "SharePointGroup")
                    {
                        SP.Group groupMembers = _context.Web.SiteGroups.GetByName(ra.Member.Title);
                        _context.Load(groupMembers, group => group.Users);
                        await _context.ExecuteQueryWithIncrementalRetry();

                        foreach(SP.User usr in groupMembers.Users)
                        {
                            sb.Add(usr.LoginName);
                            jobLog.Write("  " + usr.LoginName + " : " + usr.PrincipalType);
                        }

                    }
                    else
                    {
                        if ((ra.Member.PrincipalType.ToString() == "SecurityGroup") || (ra.Member.PrincipalType.ToString() == "User"))
                        {
                            sb.Add(ra.Member.LoginName);
                        }
                    } 
                }
                return sb;
            }
            catch(System.Exception e)
            {
                jobLog.Write("ユーザー情報を取得できませんでした。" + e);
                return null;
            }
        }

        #region 今回利用しない
        /// <summary>
        /// 対象サイトに権限があるユーザー全てを取得する。
        /// </summary>
        /// <returns></returns>
        public async Task<SP.ListItemCollection> GetAuthorizedUsers()
        {
            jobLog.Write("ユーザー情報を取得します。");
            try
            {
                SP.List l = _context.Web.SiteUserInfoList;
                _context.Load(l);
                await _context.ExecuteQueryWithIncrementalRetry();

                // ユーザーとセキュリティグループのみを取得
                SP.CamlQuery query = new SP.CamlQuery();
                query.ViewXml = @"<View>
                                    <Query>
                                        <Where>
                                            <Or>
                                                <Contains>
                                                    <FieldRef Name='Name'/>
                                                    <Value Type='Text'>|tenant|</Value>
                                                </Contains>
                                                <Contains>
                                                    <FieldRef Name='Name'/>
                                                    <Value Type='Text'>|membership|</Value>
                                                </Contains>
                                            </Or>
                                        </Where>
                                    </Query>
                                </View>";

                SP.ListItemCollection collListItem = l.GetItems(query);
                _context.Load(collListItem);
                await _context.ExecuteQueryWithIncrementalRetry();

                foreach (SP.ListItem item in collListItem)
                {
                    jobLog.Write("Name:" + item["Name"] + "ID:" + item.Id);
                }
                return collListItem;

            }

            catch (System.Exception e)
            {
                jobLog.Write("ユーザー情報を取得できませんでした。" + e);
                return null;
            }
        }
        #endregion

        /// <summary>
        /// 対象のSPグループからユーザーを全削除する。
        /// </summary>
        /// <param name="SPG">SPグループ名を指定</param>
        /// <returns></returns>
        public async Task<bool> RemoveTargetSPGroupUsers(string SPG)
        {
            jobLog.Write("=== SPG内のユーザーを削除します。=> " + SPG + " ===");

            try
            {

                var groups = _context.Web.SiteGroups;
                _context.Load(groups);
                await _context.ExecuteQueryAsync();

                SP.Group SPGroups = groups.GetByName(SPG);
                _context.Load(SPGroups.Users);
                await _context.ExecuteQueryWithIncrementalRetry();

                try
                {
                    foreach (SP.User user in SPGroups.Users)
                    {
                        SPGroups.Users.RemoveByLoginName(user.LoginName);
                        jobLog.Write("Remove " + user.LoginName.ToString());
                    }
                }
                catch (System.AggregateException e)
                {
                    Console.WriteLine(e);
                }
                SPGroups.Update();
                _context.Web.Update();
                await _context.ExecuteQueryWithIncrementalRetry();

                return true;

            }
            catch(System.Exception e)
            {
                jobLog.Write("=== SPG内のユーザーを取得できませんでした。" + e);
                return false;
            }
        }

        /// <summary>
        /// ユーザー追加
        /// </summary>
        /// <param name="spg"></param>
        /// <param name="users"></param>
        /// <returns></returns>
        public async Task<bool> AddUsersToGroup(string spg, List<string> users)
        {
            jobLog.Write("=== SPG内にユーザーを登録します。=> " + spg + " ===");

            SP.Group SPGroup = _context.Web.SiteGroups.GetByName(spg);
            try
            {
                _context.Load(SPGroup);
                await _context.ExecuteQueryWithIncrementalRetry();
            }
            catch (System.Exception e)
            {
                jobLog.Write("SPGを取得できませんでした。。\r\n" + e);
                return false;
            }

            //var ensuredUsers = users.Select(x => _context.Web.EnsureUser(x));
            //ensuredUsers.ToList().ForEach(x => _context.Load(x));
            //await _context.ExecuteQueryWithIncrementalRetry();

            foreach (string usr in users)
            {
                jobLog.Write(usr);
                SP.User user = _context.Web.EnsureUser(usr);
                SPGroup.Users.AddUser(user);
                try
                {
                    SPGroup.Update();
                    await _context.ExecuteQueryWithIncrementalRetry();
                }
                catch (System.Exception e)
                {
                    jobLog.Write("SPG内にユーザーを登録できませんでした。\r\n" + e);
                    continue;
                }
                //SPGroup.Update();
                //await _context.ExecuteQueryWithIncrementalRetry();
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

    #region ExecuteQueryを実行する際に挟むメソッドSPの「調整」対策
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
    #endregion

}
