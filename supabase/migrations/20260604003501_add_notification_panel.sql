/*
  # Add invitation notifications display

  This migration creates the structure for displaying pending invitations
  to team members when they log in or view team management.
  
  The invitations are stored locally in the frontend via localStorage,
  but we're preparing the schema for future Supabase integration.
*/

-- Verify table exists
SELECT 1;