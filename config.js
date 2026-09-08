window.APP_CONFIG = {
  // OAuth Client ID tipa "Web application" iz istega Google Cloud projekta kot Apps Script.
  OAUTH_CLIENT_ID: 'VSTAVI_OAUTH_CLIENT_ID.apps.googleusercontent.com',

  // Deployment ID API executable (NI /exec URL in NI Script ID).
  API_DEPLOYMENT_ID: 'VSTAVI_API_EXECUTABLE_DEPLOYMENT_ID',

  // Obseg, ki ga uporablja trenutni Code.gs.
  OAUTH_SCOPES: 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/userinfo.email',

  ALLOWED_USERS: [
    'joze.malnaric@gmail.com',
    's.malnaric@gmail.com',
    'helena.malnaric@gmail.com',
    'jozemalnaric.semic@gmail.com'
  ]
};
