/* Sallee Event Calendar — settings. Fill these in once (see README.md, step 2).
   Leave clientId empty to run the page in demo mode with the sample events. */
window.SALLEE_CAL_CONFIG = {
  // Entra ID app registration (can be the same one the dispatch board uses)
  clientId: "",                 // Application (client) ID
  tenantId: "",                 // Directory (tenant) ID

  // SharePoint site that holds the data
  siteHost: "YOURTENANT.sharepoint.com",   // e.g. salleehorsevans.sharepoint.com
  sitePath: "/sites/SocialMediaMarketing", // the part after the host in the site's URL

  // List names (created for you on first run)
  eventsList: "Event Calendar",
  recipientsList: "Event Reminder Recipients",

  // How often the page checks SharePoint for other people's changes
  pollSeconds: 60
};
