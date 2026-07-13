# Email Setup Guide - Team Invitations

This application sends invitation emails to team members using **Resend**, a modern email service.

## How to Enable Email Invitations

### Step 1: Get a Resend API Key (Free)

1. Visit [https://resend.com](https://resend.com)
2. Sign up for a free account
3. Go to the API Keys section
4. Copy your API key (starts with `re_`)
5. Free tier includes 3,000 emails/month

### Step 2: Add the API Key to Supabase

1. Go to your Supabase Project Dashboard
2. Navigate to **Edge Functions** section
3. Click on **Secrets** (or environment variables)
4. Create a new secret:
   - **Name**: `RESEND_API_KEY`
   - **Value**: Paste your Resend API key from Step 1
5. Save the secret

### Step 3: Test Email Sending

1. Navigate to **Team Management** page
2. Create a team (if you don't have one)
3. Click "Add Member"
4. Enter a team member's email and name
5. Click "Add Member"
6. Check the email inbox - the invitation should arrive within seconds

## How It Works

The invitation link is generated automatically from the current browser URL
(`window.location.origin`), so it always points to the correct deployed domain —
no manual `APP_URL` configuration required.

When `RESEND_API_KEY` is not configured:
- Invitations are still created in the database
- The accept-invitation link is generated but no email is sent
- A warning is logged in the browser console
- The edge function returns `emailSent: false`

When the key is present:
- A professionally formatted HTML email is sent via Resend
- The response includes `emailSent: true` and the Resend message ID

## What Members Receive

Team members receive a professionally formatted HTML email containing:
- Team name and inviter information
- List of team features and benefits
- A clickable "Accept Invitation" button
- Backup link if button doesn't work

## Troubleshooting

### "Email service not configured" message
This means the `RESEND_API_KEY` secret hasn't been added to Supabase yet. Follow Step 2 above.

### Email not arriving
1. Check spam/junk folder
2. Verify the email address is correct
3. Ensure your Resend account is verified for sending
4. Check Resend dashboard for delivery status

### Invalid email format
Resend requires a valid sender domain. By default, this app uses `noreply@launchpad-suite.com`. For production:
1. Verify your domain with Resend
2. Update the sender email in the edge function (`SENDER_EMAIL` constant)

## Production Considerations

For production deployment:
1. Configure a custom sender domain in Resend
2. Add custom branding/templates
3. Monitor email delivery in Resend dashboard
4. Implement email unsubscribe links for compliance
5. Set up DKIM, SPF, and DMARC records for your domain
