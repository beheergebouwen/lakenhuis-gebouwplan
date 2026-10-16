// Lakenhuis Gebouwplan – instellingen (door IT in te vullen, zie HANDLEIDING.md)
window.LK_CONFIG = {
  tenantId: "a3db1c8a-cea2-419f-8f9b-3726382030ab",                  // Entra ID → Overzicht → Tenant-id
  clientId: "52337422-ab05-42d3-a53a-ad3e043cef7e",      // App-registratie "Lakenhuis Gebouwplan" → Application (client) ID
  siteHostname: "overmolen.sharepoint.com",      // controleer: <naam>.sharepoint.com
  sitePath: "/sites/OVM_DATA",                // SharePoint-site waar de lijsten komen
  planFolder: "Lakenhuis/Plannen",               // map in de documentbibliotheek van die site
  lists: {
    rooms: "LK_Ruimten", doors: "LK_Deuren", meters: "LK_IstaMeters",
    orgs: "LK_Organisaties", settings: "LK_Instellingen"
  }
};
