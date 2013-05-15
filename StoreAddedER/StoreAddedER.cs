using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using System.Diagnostics;
using System.Text;
using Microsoft.SharePoint.Client;
using System.Net;
using System.Collections.Generic;

namespace UPCOR.CreateGroupWhenStore.StoreAddedER
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class StoreAddedER : SPItemEventReceiver
    {
        private EventLog _log = null;
        private const string _source = "UPCOR.CreateGroupWhenStore";
        private const string _delim = ": ";
        private StringBuilder sbDebug = new StringBuilder();

        public EventLog Log {
            get {
                if (_log == null) {
                    if (!EventLog.SourceExists(_source))
                        EventLog.CreateEventSource(_source, "Application");
                    _log = new EventLog();
                    _log.Source = _source;
                }
                return _log;
            }
        }

        /// <summary>
        /// An item is being updated
        /// </summary>
        public override void ItemUpdating(SPItemEventProperties properties) {
            base.ItemUpdating(properties);
            sbDebug.AppendLine("ItemUpdating");

            bool hasKundnummer = false;
            bool kundnummerChanged = false;

            if (properties.WebUrl != "http://web1.upcor.se/sites/blg")
                return;
            try {
                #region skriver ut debuginformation, kollar om kundnumret är ändrat
                sbDebug.AppendLine("StoreAddedER ItemUpdating");
                StringBuilder sb = new StringBuilder("StoreAddedER ItemUpdating");
                sb.AppendLine();
                sb.AppendLine();
                sb.AppendLine("ErrorMessage" + _delim + properties.ErrorMessage);
                sb.AppendLine("EventType" + _delim + properties.EventType.ToString());
                sb.AppendLine("ListId" + _delim + properties.ListId.ToString());
                sb.AppendLine("ListItemId" + _delim + properties.ListItemId.ToString());
                sb.AppendLine("ListTitle" + _delim + properties.ListTitle.ToString());
                sb.AppendLine("SiteId" + _delim + properties.SiteId.ToString());
                sb.AppendLine("Status" + _delim + properties.Status.ToString());
                sb.AppendLine("UserDisplayName" + _delim + properties.UserDisplayName.ToString());
                sb.AppendLine("UserLoginName" + _delim + properties.UserLoginName.ToString());
                sb.AppendLine("WebUrl" + _delim + properties.WebUrl.ToString());
                sb.AppendLine("ContentTypeId" + _delim + properties.ListItem.ContentTypeId.ToString());
                sbDebug.AppendLine("pre-Kundnummer");
                try {
                    if (properties.ListItem.Fields.ContainsField("Kundnummer")) {
                        string knr = (string)properties.ListItem["Kundnummer"];
                        if (knr == null) {
                            sbDebug.AppendLine("Kundnummer null");
                        }
                        else if (knr.Trim() == string.Empty) {
                            sbDebug.AppendLine("Kundnummer whitespace/empty");
                        }
                        else {
                            sb.AppendLine("Kundnummer" + _delim + knr.ToString());
                            hasKundnummer = true;
                        }
                    }
                }
                catch (Exception kex) {
                    sbDebug.AppendLine("Kundnummer Exception: " + kex.Message);
                    sbDebug.AppendLine("Kundnummer Stacktrace: " + kex.StackTrace);
                }
                //sbDebug.AppendLine("pre-Fields");
                //foreach (SPField f in properties.ListItem.Fields) {
                //    sb.AppendLine("Fields - Key: " + f.Title + " - Value: " + properties.ListItem[f.Id]);
                //}
                sbDebug.AppendLine("pre-AfterProperties");
                string forraKundnumret = string.Empty;
                foreach (System.Collections.DictionaryEntry p2 in properties.AfterProperties) {
                    string before = properties.ListItem.GetFormattedValue(p2.Key.ToString());
                    sb.AppendLine();
                    sb.AppendLine("Field - " + p2.Key);
                    sb.AppendLine("Before: " + before);
                    sb.AppendLine("After: " + p2.Value);
                    if ((string)p2.Key == "Kundnummer") {
                        forraKundnumret = before;
                        if (forraKundnumret != p2.Value) {
                            kundnummerChanged = true;
                        }
                    }
                }
                sb.AppendLine("KundnummerChanged" + _delim + kundnummerChanged.ToString());
                Log.WriteEntry(sb.ToString(), EventLogEntryType.Information, 1000);
#endregion

                //if (properties.ListId == new Guid("4b9e43d3-0a09-47b9-8e20-4ad11d413759")) {
                if (properties.ListTitle == "Försäljningsställen" &&
                    hasKundnummer &&
                    kundnummerChanged) {
                    #region Hämtar det nya kundnumret
                    sbDebug.AppendLine("Försäljningsställen && hasKundnummer && kundnummerChanged");
                    string kundnummer = null;
                    foreach (System.Collections.DictionaryEntry ap in properties.AfterProperties) {
                        if ((string)ap.Key == "Kundnummer") {
                            kundnummer = (string)ap.Value;
                            break;
                        }
                    }
                    #endregion
                    #region Hämtar det gamla kundnumret om det inte är ändrat
                    sbDebug.AppendLine("Kolla om kundnummer redan är inskrivet");
                    if (string.IsNullOrWhiteSpace(kundnummer)) {
                        kundnummer = (string)properties.ListItem["Kundnummer"];
                    }
                    #endregion
                    if (!string.IsNullOrWhiteSpace(kundnummer)) {
                        sbDebug.AppendLine("!string.IsNullOrWhiteSpace(kundnummer)");
                        // Inloggad användare
                        SPMember member = properties.Web.AllUsers.GetByID(properties.CurrentUserId);
                        int groupId = 0;
                        try {
                            SPSecurity.RunWithElevatedPrivileges(delegate() {
                                using (SPSite elevatedSite = new SPSite(properties.SiteId)) {
                                    SPWeb elevatedWeb = elevatedSite.RootWeb;

                                    #region Leta upp grupp med kundnummer, skapa om den inte finns
                                    sbDebug.AppendLine("Leta upp grupp med kundnummer, skapa om den inte finns");

                                    SPGroup group = null;

                                    //var groups = properties.Web.SiteGroups.GetCollection(new string[] { kundnummer });
                                    var groups = elevatedWeb.SiteGroups.GetCollection(new string[] { kundnummer });
                                    if (groups.Count == 0) {
                                        //properties.Web.SiteGroups.Add(kundnummer, member, null, string.Empty);
                                        //group = properties.Web.SiteGroups.GetByName(kundnummer);
                                        elevatedWeb.SiteGroups.Add(kundnummer, member, null, string.Empty);
                                        group = elevatedWeb.SiteGroups.GetByName(kundnummer);
                                    }
                                    else {
                                        group = groups[0];
                                    }


                                    SPRoleAssignment assignment = null;

                                    if (group != null) {
                                        groupId = group.ID;
                                        assignment = new SPRoleAssignment(group);
                                    }

                                    #endregion

                                    #region Leta upp behörigheten för att redigera
                                    sbDebug.AppendLine("Leta upp behörigheten för att redigera");
                                    SPRoleDefinition role = null;
                                    bool done = false;
                                    foreach (SPRoleDefinition rd in elevatedWeb.RoleDefinitions) {
                                        if (done)
                                            break;
                                        switch (rd.Name.ToLower()) {
                                            //case "fullständig behörighet":
                                            //case "full control":
                                            case "redigera":
                                            case "edit":
                                                role = rd;
                                                done = true;
                                                break;
                                        }
                                    }
                                    #endregion

                                    #region Ge redigerarättigheter för kundnummer-gruppen till det ändrade objektet
                                    sbDebug.AppendLine("Lägg assignment till role");
                                    assignment.RoleDefinitionBindings.Add(role);
                                    sbDebug.AppendLine("Hämtar lista");
                                    SPList currentList = elevatedWeb.Lists.GetList(properties.ListId, false);
                                    sbDebug.AppendLine("Hämtar listitem");
                                    SPListItem liChanged = currentList.Items.GetItemById(properties.ListItemId);
                                    sbDebug.AppendLine("-  Redigera till det ändrade objektet");
                                    sbDebug.AppendLine("   bryter arv");
                                    liChanged.BreakRoleInheritance(true);
                                    sbDebug.AppendLine("   sätter assignment");
                                    liChanged.RoleAssignments.Add(assignment);
                                    #endregion
                                }
                            });
                        }
                        catch (Exception ex) {
                            #region Logga exception
                            StringBuilder sbErr = new StringBuilder();
                            sbErr.AppendLine("Misslyckades att skapa grupp!");
                            sbErr.AppendLine();
                            sbErr.AppendLine("Debug: ");
                            sbErr.AppendLine(sbDebug.ToString());
                            sbErr.AppendLine();
                            sbErr.AppendLine("Message: " + ex.Message);
                            sbErr.AppendLine("Stacktrace: " + ex.StackTrace);
                            Log.WriteEntry(sbErr.ToString(), EventLogEntryType.Error, 1102);
                            #endregion
                        }

                        if (groupId > 0) {
                            Guid siteid = properties.Site.ID;
                            Guid webid = properties.Web.ID;

                            #region Hämtar IDn på adress, ägare och kontakter
                            sbDebug.AppendLine("Ge rättigheter");
                            sbDebug.AppendLine("Grupp: " + groupId.ToString());

                            string adress = null;
                            try {
                                //adress = properties.AfterProperties["Adress"].ToString();
                                adress = (string)properties.ListItem["Adress"];
                                adress = adress.Substring(0, adress.IndexOf(';'));
                                sbDebug.AppendLine("Adress: " + adress.ToString());
                            }
                            catch (Exception aex) {
                                sbDebug.AppendLine("Adress Exception: " + aex.Message);
                                sbDebug.AppendLine("Adress Stacktrace: " + aex.StackTrace);
                            }

                            string agare = null;
                            try {
                                //agare = properties.AfterProperties["_x00c4_gare"].ToString();
                                agare = (string)properties.ListItem["_x00c4_gare"];
                                agare = agare.Substring(0, agare.IndexOf(';'));
                                sbDebug.AppendLine("_x00c4_gare: " + agare.ToString());
                            }
                            catch (Exception aaex) {
                                sbDebug.AppendLine("_x00c4_gare Exception: " + aaex.Message);
                                sbDebug.AppendLine("_x00c4_gare Stacktrace: " + aaex.StackTrace);
                            }

                            int[] kontakter = null;
                            try {
                                //kontakt = properties.AfterProperties["Kontaktperson"].ToString();
                                List<int> listKontakter = new List<int>();
                                SPFieldLookupValueCollection fieldcol = (SPFieldLookupValueCollection)properties.ListItem["Kontaktperson"];
                                foreach (SPFieldLookupValue field in fieldcol) {
                                    listKontakter.Add(field.LookupId);
                                    sbDebug.AppendLine("Kontaktperson: " + field.LookupId.ToString());
                                }
                                kontakter = listKontakter.ToArray();
                            }
                            catch (Exception knex) {
                                sbDebug.AppendLine("Kontaktperson Exception: " + knex.Message);
                                sbDebug.AppendLine("Kontaktperson Stacktrace: " + knex.StackTrace);
                            }
                            sbDebug.AppendLine("-  Hämtar");
                            #endregion

                            #region Ge redigera till grupp för adress, ägare och kontakter
                            ClientContext ctx = new ClientContext(properties.WebUrl);
                            ctx.Credentials = new NetworkCredential(Properties.Resources.User, Properties.Resources.Password, Properties.Resources.Domain);
                            sbDebug.AppendLine("   hämta redigerarättigheten");
                            //RoleDefinition rdEdit = ctx.Web.RoleDefinitions.GetByName("Redigera");
                            RoleDefinition rdEdit = ctx.Web.RoleDefinitions.GetByType(RoleType.Editor);
                            Group g = ctx.Web.SiteGroups.GetById(groupId);
                            RoleAssignmentCollection assignments = ctx.Web.RoleAssignments;

                            #region -  Redigera till Kontakter
                            if (kontakter == null) {
                                sbDebug.AppendLine("-  Skippar Kontakter");
                            }
                            else {
                                sbDebug.AppendLine("-  Redigera till Kontakter");
                                List listContacts = ctx.Web.Lists.GetByTitle("Kontakter");
                                foreach (int kontakt in kontakter) {
                                    RoleDefinitionBindingCollection rdb = new RoleDefinitionBindingCollection(ctx);
                                    rdb.Add(rdEdit);
                                    ListItem li = listContacts.GetItemById(kontakt);
                                    //li.ResetRoleInheritance();
                                    li.BreakRoleInheritance(true, true);
                                    li.RoleAssignments.Add(g, rdb);
                                    li.Update();
                                    //ctx.Load(li.RoleAssignments);
                                    //sbDebug.AppendLine("Kör query(1a)");
                                    //ctx.ExecuteQuery();
                                    //foreach (RoleAssignment ra in li.RoleAssignments) {
                                    //    ctx.Load(ra.Member);
                                    //    ctx.Load(ra.RoleDefinitionBindings);
                                    //    ctx.ExecuteQuery();
                                    //    sbDebug.AppendLine("Kontakt ra: " + ra.Member.LoginName);
                                    //    foreach(var binding in ra.RoleDefinitionBindings) {
                                    //        sbDebug.AppendLine(binding.Name);
                                    //    }
                                    //}
                                    //sbDebug.AppendLine();
                                }

                            }
                            #endregion

                            #region -  Redigera till Ägare
                            if (string.IsNullOrWhiteSpace(agare)) {
                                sbDebug.AppendLine("-  Skippar Ägare");
                            }
                            else {
                                RoleDefinitionBindingCollection rdb = new RoleDefinitionBindingCollection(ctx);
                                rdb.Add(rdEdit);
                                sbDebug.AppendLine("-  Redigera till Ägare");
                                sbDebug.AppendLine("rdbc count(3b) " + rdb.Count.ToString());
                                List listOwners = ctx.Web.Lists.GetByTitle("Ägare");
                                ListItem li = listOwners.GetItemById(agare);
                                //li.ResetRoleInheritance();
                                li.BreakRoleInheritance(true, true);
                                li.RoleAssignments.Add(g, rdb);
                                li.Update();
                                //sbDebug.AppendLine("Kör query(1b)");
                                //ctx.ExecuteQuery();
                            }
                            #endregion

                            #region -  Redigera till Försäljningsställen Adresser
                            if (string.IsNullOrWhiteSpace(adress)) {
                                sbDebug.AppendLine("-  Skippar Försäljningsställen Adresser");
                            }
                            else {
                                RoleDefinitionBindingCollection rdb = new RoleDefinitionBindingCollection(ctx);
                                rdb.Add(rdEdit);
                                sbDebug.AppendLine("-  Redigera till Försäljningsställen Adresser");
                                List listAddresses = ctx.Web.Lists.GetByTitle("Försäljningsställen Adresser");
                                ListItem li = listAddresses.GetItemById(adress);
                                //li.ResetRoleInheritance();
                                li.BreakRoleInheritance(true, true);
                                li.RoleAssignments.Add(g, rdb);
                                li.Update();
                                //sbDebug.AppendLine("Kör query(1c)");
                                //ctx.ExecuteQuery();
                            }
                            #endregion

                            sbDebug.AppendLine("Kör query(1d)");
                            ctx.ExecuteQuery();
                            #endregion

                            #region Koppla grupp till försäljningsställe
                            try {

                                sbDebug.AppendLine("Koppla grupp till försäljningsställe");
                                List listConnect = ctx.Web.Lists.GetByTitle("Grupper för försäljningsställen");
                                CamlQuery cq = new CamlQuery();
                                //                                cq.ViewXml = string.Concat(
                                //"<View><Query><Where><Eq>",
                                //    "<FieldRef Name='F_x00f6_rs_x00e4_ljningsst_x00e4' LookupId='True'/>",
                                //    "<Value Type='Lookup'>",
                                //        properties.ListItemId.ToString(),
                                //    "</Value>",
                                //"<Eq></Where></Query></View>");

                                cq.ViewXml = @"<View>  
        <Query> 
            <Where><Eq><FieldRef Name='F_x00f6_rs_x00e4_ljningsst_x00e4' LookupId='True' /><Value Type='Lookup'>" + properties.ListItemId.ToString() + @"</Value></Eq></Where> 
        </Query> 
    </View>";


                                //                                    cq.ViewXml = @"<View>  
                                //            <Query> 
                                //               <Where><Eq><FieldRef Name='Butik' LookupId='True' /><Value Type='Lookup'>" + properties.ListItemId.ToString() + @"</Value></Eq></Where> 
                                //            </Query> 
                                //      </View>";
                                sbDebug.AppendLine("Query -> " + cq.ViewXml);
                                ListItemCollection items = listConnect.GetItems(cq);
                                ctx.Load(items);
                                sbDebug.AppendLine("Kör query(2)");
                                ctx.ExecuteQuery();
                                if (items.Count == 0) {
                                    sbDebug.AppendLine("0 items i grupper för försäljningsställen");
                                    ListItem item = listConnect.AddItem(new ListItemCreationInformation { });
                                    item["F_x00f6_rs_x00e4_ljningsst_x00e4"] = properties.ListItemId;
                                    item["Grupp"] = groupId;
                                    item.Update();
                                }
                                else {
                                    sbDebug.AppendLine("Minst 1 item i grupper för försäljningsställen");
                                    foreach (ListItem item in items) {
                                        item["Grupp"] = groupId;
                                        item.Update();
                                    }
                                }

                            }
                            catch (Exception ex3) {
                                StringBuilder sb3 = new StringBuilder("StoreAddedER ItemUpdated Exception(Koppla grupp till försäljningsställe)");
                                sb3.AppendLine();
                                sb3.AppendLine("Debug: ");
                                sb3.AppendLine(sbDebug.ToString());
                                sb3.AppendLine();
                                sb3.AppendLine("Message: " + ex3.Message);
                                sb3.AppendLine("Stacktrace: " + ex3.StackTrace);
                                sb3.AppendLine("Type: " + ex3.GetType().ToString());
                                sb3.AppendLine("Data: ");
                                foreach (var d in ex3.Data.Keys) {
                                    sb3.AppendLine(d + " : " + ex3.Data[d]);
                                }
                                if (ex3.InnerException != null) {
                                    sb3.AppendLine("Inner Message: " + ex3.InnerException.Message);
                                    sb3.AppendLine("Inner Stacktrace: " + ex3.InnerException.StackTrace);
                                }

                                Log.WriteEntry(sb3.ToString(), EventLogEntryType.Error, 1103);
                            }
                            #endregion

                            sbDebug.AppendLine("Kör query(3)");
                            ctx.ExecuteQuery();
                        }
                        else {
                            sbDebug.AppendLine("GruppID <= 0)");
                        }
                        //} // if (role != null && assignment != null)

                        Log.WriteEntry(sbDebug.ToString(), EventLogEntryType.Information, 1002);
                    } // if (null or whitespace (kundnummer))
                } // if (properties.ListId == new Guid("4b9e43d3-0a09-47b9-8e20-4ad11d413759")
            } // try
            catch (Exception ex) {
                #region Logga exception
                StringBuilder sb = new StringBuilder("StoreAddedER ItemUpdated Exception(outer)");
                sb.AppendLine();
                sb.AppendLine("Debug: ");
                sb.AppendLine(sbDebug.ToString());
                sb.AppendLine();
                sb.AppendLine("Message: " + ex.Message);
                sb.AppendLine("Stacktrace: " + ex.StackTrace);
                sb.AppendLine("Type: " + ex.GetType().ToString());

                Log.WriteEntry(sb.ToString(), EventLogEntryType.Error, 1100);
                #endregion
            }
        } // ItemUpdating
    } // class StoreAddedER
} // Namespace