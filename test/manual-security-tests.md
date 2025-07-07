# Manual Security Tests

These steps verify that unauthorized or forged requests to the `process_excel` AJAX action are rejected.

1. **Logged-out CSRF attempt**
   - Open a private browser window and visit `admin-ajax.php` directly:
     ```
     https://your-site.example/wp-admin/admin-ajax.php?action=process_excel
     ```
   - You should receive a JSON error response with a `nonce_failure` code.

2. **Logged-in user without `upload_files` capability**
   - Log in as a user with minimal permissions (e.g. Subscriber).
   - Attempt to upload a file using the plugin interface or by crafting a POST request to `admin-ajax.php?action=process_excel`.
   - The request should fail with an `insufficient_permissions` error.

3. **Invalid nonce value**
   - While logged in with sufficient permissions, intercept the request made by the plugin and modify the `nonce` parameter to an invalid value.
   - Resend the request. A `nonce_failure` response should be returned.

Successful completion of the above steps confirms that the security checks are active.
