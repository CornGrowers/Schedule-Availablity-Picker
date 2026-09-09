// ---------------------------------------------------------------------------
// FILL THESE IN after registering the Azure AD app (see ADMIN_DEPLOYMENT.md
// for the organization-wide rollout, or README.md for the personal-account
// self-service version).
// ---------------------------------------------------------------------------
const APP_CONFIG = {
  // Azure AD "Application (client) ID" from the app registration.
  clientId: "REPLACE_WITH_AZURE_AD_APPLICATION_CLIENT_ID",

  // "organizations" = any work/school Microsoft Entra tenant (no personal
  // accounts). Use the org's own tenant ID here instead once known, to
  // restrict sign-in to just that tenant.
  authority: "https://login.microsoftonline.com/organizations",

  // Must exactly match a Redirect URI (type: Single-page application)
  // configured on the Azure AD app registration.
  redirectUri: "https://REPLACE_WITH_YOUR_HOSTED_DOMAIN/taskpane.html",

  // Calendars.Read: read free/busy + events (also covers checking colleagues'
  // free/busy via getSchedule). People.Read: search the org directory when
  // adding colleagues.
  graphScopes: ["Calendars.Read", "People.Read"],

  graphBaseUrl: "https://graph.microsoft.com/v1.0",
};
