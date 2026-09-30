/* Sallee Event Calendar — settings for Sallee's Microsoft 365 tenant.
   Leave clientId empty to run the page in demo mode with the sample events. */
window.SALLEE_CAL_CONFIG = {
  // Entra ID app registration: "Sallee Event Calendar"
  clientId: "3b8c2e35-6037-4945-8fa4-30e6da1c8b30",   // Application (client) ID
  tenantId: "428e82a0-5a1b-41e6-a24c-ca8ea70db05a",   // Directory (tenant) ID

  // SharePoint site that holds the data
  siteHost: "salleehorsevans.sharepoint.com",
  sitePath: "/sites/SOCIALMEDIAANDMARKETING",

  // List names (created for you on first run)
  eventsList: "Event Calendar",
  recipientsList: "Event Reminder Recipients",

  // How often the page checks SharePoint for other people's changes
  pollSeconds: 60
};
