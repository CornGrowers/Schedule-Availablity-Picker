// ---------------------------------------------------------------------------
// FILL THESE IN after registering the Azure AD app (see ADMIN_DEPLOYMENT.md
// for the organization-wide rollout, or README.md for the personal-account
// self-service version).
// ---------------------------------------------------------------------------
const APP_CONFIG = {
  // Azure AD "Application (client) ID" from the app registration.
  clientId: "e2f145fa-efec-4218-baeb-9b15c04b5668",

  // "organizations" = any work/school Microsoft Entra tenant (no personal
  // accounts). Use the org's own tenant ID here instead once known, to
  // restrict sign-in to just that tenant.
  authority: "https://login.microsoftonline.com/cedfb7f7-0481-4eda-a915-80c712b7e9e0",

  // Must exactly match a Redirect URI (type: Single-page application)
  // configured on the Azure AD app registration.
  redirectUri: "https://corngrowers.github.io/Schedule-Availablity-Picker/taskpane.html",

  // Calendars.Read: read free/busy + events (also covers checking colleagues'
  // free/busy via getSchedule). People.Read: search the org directory when
  // adding colleagues.
  graphScopes: ["Calendars.Read", "People.Read"],

  graphBaseUrl: "https://graph.microsoft.com/v1.0",
};
