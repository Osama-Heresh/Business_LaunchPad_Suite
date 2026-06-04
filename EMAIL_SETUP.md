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

1. Start the development server
2. Navigate to **Team Management** page
3. Create a team (if you don't have one)
4. Click "Add Member"
5. Enter a team member's email and name
6. Click "Add Member"
7. Check the email inbox - the invitation should arrive within seconds

## What Members Receive

Team members will receive a professionally formatted HTML email containing:
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
2. Update the sender email in the edge function

## Production Considerations

For production deployment:
1. Configure a custom sender domain in Resend
2. Add custom branding/templates
3. Monitor email delivery in Resend dashboard
4. Implement email unsubscribe links for compliance
5. Set up DKIM, SPF, and DMARC records for your domain

## Current Status

- ✓ Edge function deployed and ready
- ⏳ Waiting for RESEND_API_KEY to be added to Supabase
- Once key is added, emails will send automatically
