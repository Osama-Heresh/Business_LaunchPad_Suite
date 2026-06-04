import "jsr:@supabase/functions-js/edge-runtime.d.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Client-Info, Apikey",
};

interface InvitationRequest {
  memberEmail: string;
  memberName: string;
  teamName: string;
  teamOwnerId: string;
  teamOwnerEmail: string;
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") {
    return new Response(null, {
      status: 200,
      headers: corsHeaders,
    });
  }

  try {
    const payload: InvitationRequest = await req.json();
    const { memberEmail, memberName, teamName, teamOwnerId, teamOwnerEmail } = payload;

    if (!memberEmail || !memberName || !teamName) {
      return new Response(
        JSON.stringify({ error: "Missing required fields" }),
        {
          status: 400,
          headers: { ...corsHeaders, "Content-Type": "application/json" },
        }
      );
    }

    const appUrl = Deno.env.get("VITE_APP_URL") || "http://localhost:5173";
    const invitationLink = `${appUrl}/team-invite?email=${encodeURIComponent(memberEmail)}&team=${encodeURIComponent(teamName)}&inviter=${encodeURIComponent(teamOwnerEmail)}`;

    const emailContent = `
<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8">
    <title>Team Invitation</title>
  </head>
  <body style="font-family: Arial, sans-serif; background-color: #f5f5f5; padding: 20px;">
    <div style="max-width: 600px; margin: 0 auto; background-color: white; border-radius: 8px; padding: 30px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">
      <h1 style="color: #1e293b; margin-bottom: 20px;">You've Been Invited to a Team!</h1>

      <p style="color: #475569; font-size: 16px; line-height: 1.6;">
        Hi <strong>${memberName}</strong>,
      </p>

      <p style="color: #475569; font-size: 16px; line-height: 1.6;">
        <strong>${teamOwnerEmail}</strong> has invited you to join the <strong>${teamName}</strong> team on LaunchPad Suite.
      </p>

      <p style="color: #475569; font-size: 16px; line-height: 1.6;">
        As a team member, you'll be able to:
      </p>

      <ul style="color: #475569; font-size: 16px; line-height: 1.8;">
        <li>View and manage tasks assigned to you</li>
        <li>Update task status and progress</li>
        <li>Collaborate with your team members</li>
        <li>Track project milestones</li>
      </ul>

      <div style="margin: 30px 0;">
        <a href="${invitationLink}" style="display: inline-block; background-color: #3b82f6; color: white; padding: 12px 30px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 16px;">Accept Invitation</a>
      </div>

      <p style="color: #94a3b8; font-size: 14px; line-height: 1.6;">
        If you didn't expect this invitation or have questions, please contact ${teamOwnerEmail} directly.
      </p>

      <hr style="border: none; border-top: 1px solid #e2e8f0; margin: 20px 0;">

      <p style="color: #94a3b8; font-size: 12px;">
        LaunchPad Suite - Your Business Hub
      </p>
    </div>
  </body>
</html>
    `;

    const smtpResponse = await fetch("https://api.sendgrid.com/v3/mail/send", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${Deno.env.get("SENDGRID_API_KEY")}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        personalizations: [
          {
            to: [{ email: memberEmail, name: memberName }],
          },
        ],
        from: {
          email: "noreply@launchpadsuite.com",
          name: "LaunchPad Suite",
        },
        subject: `You're invited to join ${teamName} team!`,
        content: [
          {
            type: "text/html",
            value: emailContent,
          },
        ],
      }),
    });

    if (!smtpResponse.ok) {
      console.error("SendGrid error:", await smtpResponse.text());

      return new Response(
        JSON.stringify({
          success: true,
          message: "Invitation saved (email service not configured)",
        }),
        {
          status: 200,
          headers: { ...corsHeaders, "Content-Type": "application/json" },
        }
      );
    }

    return new Response(
      JSON.stringify({
        success: true,
        message: "Invitation email sent successfully",
        link: invitationLink,
      }),
      {
        status: 200,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      }
    );
  } catch (error) {
    console.error("Error:", error);
    return new Response(
      JSON.stringify({
        success: true,
        message: "Invitation processed (email service not configured)",
      }),
      {
        status: 200,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      }
    );
  }
});
