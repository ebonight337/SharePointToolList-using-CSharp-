using SP = Microsoft.SharePoint.Client;

public async Task<List<string>> GetSPWebRoleAssignments(string weburl, string username, string password)
{
    Console.WriteLine("=== サイトに権限が付与されている全ユーザーを取得します。 => " + weburl +" ===");
    SP.Web RootWeb = _context.Web;
    _context.Load(RootWeb, w => w.HasUniqueRoleAssignments, w => w.Url, w => w.ServerRelativeUrl);
    _context.Load(RootWeb.RoleAssignments);
    await _context.ExecuteQueryAsync();

    List<string> sb = new List<string>();
    foreach (SP.RoleAssignment ra in RootWeb.RoleAssignments)
    {
        _context.Load(ra.Member);
        await _context.ExecuteQueryAsync();
        Console.WriteLine(ra.Member.LoginName + " : " + ra.Member.PrincipalType);

        if (ra.Member.PrincipalType.ToString() == "SharePointGroup")
        {
            SP.Group groupMembers = _context.Web.SiteGroups.GetByName(ra.Member.Title);
            _context.Load(groupMembers, group => group.Users);
            await _context.ExecuteQueryAsync();

            foreach(SP.User usr in groupMembers.Users)
            {
                sb.Add(usr.LoginName);
                Console.WriteLine("  " + usr.LoginName + " : " + usr.PrincipalType);
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
    Console.WriteLine("権限の有無確認を行います");
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
            Console.WriteLine("このユーザーは存在しません" + e);
            continue;
        }
    }
    return sb;
}
