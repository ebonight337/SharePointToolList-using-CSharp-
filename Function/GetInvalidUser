using System;
using System.Net;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Threading;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;

namespace GetInvalidUser
{
    /// <summary>
    /// SharePoint のデータ操作機能を提供します
    /// </summary>
    public class SharePointDataService:IDisposable
    {
        private readonly SP.ClientContext _context;
        private readonly SecureString _secureString;
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
                SP.Web RootWeb = _context.Web;
                _context.Load(RootWeb, w => w.HasUniqueRoleAssignments, w => w.Url, w => w.ServerRelativeUrl);
                _context.Load(RootWeb.RoleAssignments);
                await _context.ExecuteQueryAsync();

                List<string> sb = new List<string>();
                foreach (SP.RoleAssignment ra in RootWeb.RoleAssignments)
                {
                    _context.Load(ra.Member);
                    await _context.ExecuteQueryAsync();

                    if (ra.Member.PrincipalType.ToString() == "SharePointGroup")
                    {
                        SP.Group groupMembers = _context.Web.SiteGroups.GetByName(ra.Member.Title);
                        _context.Load(groupMembers, group => group.Users);
                        await _context.ExecuteQueryAsync();

                        foreach(SP.User usr in groupMembers.Users)
                        {
                            sb.Add(usr.LoginName);
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

                // 権限の有無確認
                jobLog.Write("権限の有無確認を行います");
                foreach (string i in sb)
                {
                    try
                    {
                        if (i.Contains("|membership|"))
                        {
                            SP.UserProfiles.PeopleManager peopleManager = new SP.UserProfiles.PeopleManager(_context);
                            var managerData = peopleManager.GetUserProfileProperties(i);
                            await _context.ExecuteQueryAsync();

                        }
                        else
                        {
                            continue;
                        }
                    }
                    catch (System.Exception e)
                    {
                            errorLog.WriteCSV(i);
                        continue;
                    }
                    
                }
                return sb;
            }
            catch(System.Exception e)
            {
                return null;
            }
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

}
