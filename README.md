# Getting automatically a miscrosoft graph token inside an http://intranet.site using a Tampermonkey Authentication Flow

## Summary:
This script enables automatic authentication for http://viajes.cdti.es (a vpn protected intranet site) using an accesible public url https://www.google.com/search?q=redirige_viajes as the OAuth REDIRECT_URI. 
Since OAuth requires https:// and does not allow http://, this workaround ensures a seamless login experience. When authentication is completed via Microsoft, the script detects the token on Google’s search page, extracts it, and redirects the user back to the intranet site viajes.cdti.es with the token included. This approach leverages Google’s universally accessible domain to bypass OAuth restrictions while maintaining security and usability. 

**Tampermonkey Authentication and Redirection Process**

### 0. **Register your app in Azure**
Summary of How to Get Them:
- TENANT_ID: Found in Azure AD > Properties > Directory ID.
- CLIENT_ID: Found in Azure AD > App registrations > Your App > Application (client) ID.
- REDIRECT_URI: Configured in Azure AD > App registrations > Your App > Authentication > Redirect URIs.
Configure REDIRECT_URI https://www.google.com/search?q=redirige_viajes

### 1. **Authentication Flow**
- The script runs on your `WEBAPP_URI`  `viajes.cdti.es` (your intranet app) 
- The script runs also on the `REDIRECT_URI` `https://www.google.com/search?q=redirige_viajes*` as the `REDIRECT_URI`.
- It first checks if the current URL matches `REDIRECT_URI`.
- If true, it extracts the authentication token from the URL fragment.
- Then, it redirects the user back to `WEBAPP_URI` (i.e., `viajes.cdti.es`) with the token.
- If already on `WEBAPP_URI`, it retrieves the stored token or initiates authentication.

### 2. **How the Redirection Works**
- Microsoft authentication sends the user to `REDIRECT_URI` after login.
- The script detects this and reconstructs the original URL for `viajes.cdti.es`.
- It appends the received token as a query parameter.
- The browser is then redirected to `viajes.cdti.es`, where the token can be used for API calls.

### 3. **Why Use Google as `REDIRECT_URI`?**
- The domain `google.com` is always accessible and avoids needing extra configuration in Azure.
- The script intercepts this page load and immediately redirects to `viajes.cdti.es`, ensuring a seamless login experience.

### 4. **How Authentication is Handled in the Script**
- The script defines `AUTH_URL` using `TENANT_ID`, `CLIENT_ID`, and `SCOPES`.
- If no valid access token is found, it redirects the user to Microsoft’s login page.
- Upon successful login, Microsoft appends the token to the `REDIRECT_URI`.
- The script extracts this token, stores it, and redirects the user to `WEBAPP_URI`.

### 5. **Handling Errors and Token Expiry**
- If an invalid or expired token is detected, the script clears it and forces re-authentication.
- Common authentication errors are handled with alerts and console logs.
- If permissions are missing, the user is notified and asked to request authorization.

### 6. **Conclusion**
- The script provides a workaround for authentication by leveraging Google’s search page as an intermediate step.
- Once authenticated, it ensures smooth redirection back to `viajes.cdti.es` with the necessary credentials.
- The implementation allows seamless and secure integration with Microsoft Graph APIs.

