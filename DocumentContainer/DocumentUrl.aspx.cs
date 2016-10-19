using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Security;
using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Web.UI.HtmlControls;
using Microsoft.SharePoint.Client.Sharing;
namespace DocumentContainer
{
    public partial class DocumentUrl : System.Web.UI.Page
    {
        #region environmental variables
        string siteUrl = ConfigurationManager.AppSettings.Get("siteUrl");
        string rootSiteUrl = ConfigurationManager.AppSettings.Get("rootSiteUrl");
        string documentLibrary = "DocumentContainer";
        #endregion

        protected void Page_Load(object sender, EventArgs e)
        {
            Uri siteUri = new Uri(siteUrl);

            string realm = TokenHelper.GetRealmFromTargetUrl(siteUri);
            string accessToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, siteUri.Authority, realm).AccessToken;
            //Creating client context
            using (ClientContext context = TokenHelper.GetClientContextWithAccessToken(siteUri.ToString(), accessToken))
            {
                //Creating web object
                Web web = context.Web;
                List list = web.Lists.GetByTitle(documentLibrary);
                context.Load(list);
                context.ExecuteQuery();

                #region Code to handle document libraries that have files > 5000

                ListItemCollectionPosition itemPosition = null;
                //This while loop works till the batches of 2000 items is completely processed
                while (true)
                {
                    CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ListItemCollectionPosition = itemPosition;
                    camlQuery.ViewXml = @"<View Scope='RecursiveAll'>
                                        <Query>
                                            <OrderBy Override='TRUE'>
                                                <FieldRef Name='ID' Ascending='TRUE' />
                                            </OrderBy>
                                        </Query>
                                        <RowLimit>2000</RowLimit>
                                     </View>";

                    Microsoft.SharePoint.Client.ListItemCollection listItems = list.GetItems(camlQuery);
                    context.Load(listItems);
                    context.ExecuteQuery();

                    itemPosition = listItems.ListItemCollectionPosition;

                    foreach (Microsoft.SharePoint.Client.ListItem item in listItems)
                    {
                        context.Load(item.File);
                        context.ExecuteQuery();

                        ObjectSharingInformation obs = ObjectSharingInformation.GetObjectSharingInformation(context, item, false, true, false, true, true, true, true);
                        context.Load(obs);
                        context.ExecuteQuery();

                        // Get file link URL
                        string fileUrl = ResolveShareUrl(context, item.File.ServerRelativeUrl);

                        // Get current sharing settings
                        if (!obs.IsSharedWithGuest)
                        {
                            // Share document for given email address
                            string sharingResult = context.Web.CreateAnonymousLinkForDocument(fileUrl, ExternalSharingDocumentOption.View);

                            HtmlAnchor anchor = new HtmlAnchor();
                            anchor.HRef = sharingResult;
                            anchor.Title = item.File.Name;
                            anchor.InnerText = item.File.Name;
                            anchor.Target = "_Parent";
                            searchedDocuments.Controls.Add(anchor);

                            searchedDocuments.Controls.Add(new LiteralControl("<br />"));

                        }
                        else
                        {
                            HtmlAnchor anchor = new HtmlAnchor();
                            anchor.HRef = context.Web.GetObjectSharingSettingsForDocument(fileUrl, true).ObjectSharingInformation.AnonymousViewLink;
                            anchor.Title = item.File.Name;
                            anchor.InnerText = item.File.Name;
                            anchor.Target = "_Parent";
                            searchedDocuments.Controls.Add(anchor);

                            searchedDocuments.Controls.Add(new LiteralControl("<br />"));
                        }
                    }

                    if (itemPosition == null)
                    {
                        break;
                    }
                }

                #endregion
            }
        }

        protected void btnGetDocument_Click(object sender, EventArgs e)
        {
        }

        private string ResolveShareUrl(ClientContext ctx, string fileServerRelativeUrl)
        {
            if (!ctx.Web.IsObjectPropertyInstantiated("Url"))
            {
                ctx.Load(ctx.Web, w => w.Url);
                ctx.ExecuteQuery();
            }
            var tenantStr = ctx.Web.Url.ToLower().Replace("-my", "").Substring(8);
            tenantStr = tenantStr.Substring(0, tenantStr.IndexOf("."));
            return String.Format("https://{0}.sharepoint.com{1}", tenantStr, fileServerRelativeUrl);
        }
    }
}