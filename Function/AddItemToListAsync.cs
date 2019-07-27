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
