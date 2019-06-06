using System;
using System.Net;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Threading;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;

namespace CreateandRemoveList
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
        public SharePointDataService(string sitePath, string mailAddress, string password) {
            _context = new SP.ClientContext(sitePath);

            _secureString = new SecureString();
            password.ToList().ForEach(c => _secureString.AppendChar(c));
            _context.Credentials = new SP.SharePointOnlineCredentials(mailAddress, _secureString);
        }


        public async Task<SP.List> CreateListAsync(string listName, string targetListTemplate, string listUrl, string Description = null) {
            Console.Write("リストの作成を実施します => " + listName);
            try {
                // 存在確認
                SP.ListCollection lists = _context.Web.Lists;
                _context.Load(lists);
                await _context.ExecuteQueryAsync();

                foreach (SP.List oList in lists) {
                    if (oList.Title == listName) {
                        Console.Write("作成しようとしているリストはすでに存在します =>" + listName);
                        throw new Exception("作成しようとしているリストはすでに存在します =>" + listName);
                    }
                }

                // リストテンプレートを取得
                SP.ListTemplateCollection templates = _context.Site.GetCustomListTemplates(_context.Web);
                _context.Load(templates);
                await _context.ExecuteQueryAsync();

                SP.ListCreationInformation listCreationInfo = new SP.ListCreationInformation();
                listCreationInfo.Title = listUrl;
                if (Description != null) {
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

            } catch (Exception e) {
                Console.Write("リストの作成中にエラーが発生しましたエラーを確認してください\r\n" + e);
                throw new Exception("リストの作成でエラーが発生しましたエラーを確認してください");
            }
            try {
                Console.Write("リストの作成が完了しました => " + listName);
                return _context.Web.Lists.GetByTitle(listName);
            } catch (Exception e) {
                Console.Write("作成したリストを取得できませんでしたエラーを確認してください\r\n" + e);
                throw new Exception("作成したリストを取得できませんでしたエラーを確認してください");
            }
        }

        public async Task<bool> DeleteListAsync(string listName) {
            Console.Write("リストの削除を実施します => " + listName);
            // 存在確認
            SP.ListCollection lists = _context.Web.Lists;
            try {
                _context.Load(lists);
                await _context.ExecuteQueryAsync();
            } catch (Exception e) {
                Console.Write("リストの取得でエラーが発生しました。\r\n" + e);
                throw new Exception("リストの取得でエラーが発生しました。エラーを確認してください。");
            }

            foreach (SP.List oList in lists) {
                if (oList.Title == listName) {
                    SP.List deleteList = _context.Web.Lists.GetByTitle(listName);
                    try {
                        deleteList.DeleteObject();
                        await _context.ExecuteQueryAsync();
                    } catch (Exception e) {
                        Console.Write("リスト削除処理で以下のエラーが発生しました\r\n" + e);
                        // throw new Exception("リスト削除処理で以下のエラーが発生しましたログを確認してください");
                    }
                    Console.Write("リストの削除が完了しました => " + listName);
                    return true;
                }
            }

            Console.Write("削除対象のリストが存在しないため、削除処理をパスします =>" + listName);
            Console.WriteLine("削除対象のリストが存在しないため、削除処理をパスします =>" + listName);
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

    }
}
