function getServiceAccountToken_() {
  const props = PropertiesService.getScriptProperties();

  const json  = JSON.parse(props.getProperty("SERVICE_ACCOUNT_JSON"));
  const email = props.getProperty("SERVICE_ACCOUNT_EMAIL");

  const now = Math.floor(Date.now() / 1000);

  const header = { alg: "RS256", typ: "JWT" };
  const claim  = {
    iss:   email,
    scope: "https://www.googleapis.com/auth/chat.bot",
    aud:   "https://oauth2.googleapis.com/token",
    exp:   now + 3600,
    iat:   now
  };

  const base64Header = Utilities.base64EncodeWebSafe(JSON.stringify(header)).replace(/=+$/, "");
  const base64Claim  = Utilities.base64EncodeWebSafe(JSON.stringify(claim)).replace(/=+$/, "");

  const unsignedJwt = base64Header + "." + base64Claim;
  const signature   = Utilities.computeRsaSha256Signature(unsignedJwt, json.private_key);
  const signedJwt   = unsignedJwt + "." + Utilities.base64EncodeWebSafe(signature).replace(/=+$/, "");

  const response = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", {
    method: "post",
    muteHttpExceptions: true,
    payload: {
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion:  signedJwt
    }
  });

  const token = JSON.parse(response.getContentText());
  if (!token.access_token) {
    throw new Error("Falha ao obter access_token: " + response.getContentText());
  }

  return token.access_token;
}
