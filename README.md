# Social Post Builder

A comprehensive social media post creation tool with AI-powered content generation and secure credential management.

## Features

- **Multi-Platform Publishing**: Support for Twitter, Facebook, Instagram, LinkedIn, TikTok, Slack, and Discord
- **Secure Credential Management**: OAuth integration with encrypted token storage
- **AI Content Generation**: 
  - Images: OpenAI DALL-E 3, Stability AI
  - Videos: Replicate Stable Video Diffusion, Runway ML
- **Rich Text Editor**: Formatting tools and platform-specific character limits
- **Text Overlay**: Customizable text overlays for images and videos
- **Scheduling**: Post immediately or schedule for later
- **Account Management**: Connect/disconnect multiple accounts per platform
- **Responsive Design**: Mobile-first, accessible interface

## Social Media Integration

### Supported Platforms
- **Twitter/X**: Full API v2 integration with media upload
- **Facebook**: Graph API with page posting capabilities
- **Instagram**: Business API for image and video posting
- **LinkedIn**: Professional network posting with media support
- **TikTok**: Video content publishing (coming soon)
- **Slack**: Workspace integration for team communication
- **Discord**: Server posting capabilities

### Security Features
- **OAuth 2.0 Authentication**: Secure token-based authentication
- **Encrypted Storage**: Sensitive credentials encrypted at rest
- **Token Refresh**: Automatic token renewal for continuous access
- **Permission Management**: Granular control over platform permissions
- **Secure Revocation**: Safe disconnection and token cleanup

## Setup

1. **Install Dependencies**
   ```bash
   npm install
   ```

2. **Configure API Keys**
   Copy `.env.example` to `.env` and add your API keys:
   ```bash
   cp .env.example .env
   ```

3. **Get API Keys**
   - **OpenAI**: https://platform.openai.com/api-keys
   - **Stability AI**: https://platform.stability.ai/account/keys
   - **Replicate**: https://replicate.com/account/api-tokens
   - **Runway ML**: https://runwayml.com/

4. **Social Media App Registration**
   Register your application with each platform:
   - **Twitter**: https://developer.twitter.com/en/portal/dashboard
   - **Facebook**: https://developers.facebook.com/apps/
   - **Instagram**: https://developers.facebook.com/docs/instagram-api/
   - **LinkedIn**: https://www.linkedin.com/developers/apps/

5. **Start Development Server**
   ```bash
   npm run dev
   ```

## Social Media API Setup

### Twitter/X API
```javascript
// Required permissions
const twitterScopes = [
  'tweet.read',
  'tweet.write',
  'users.read',
  'offline.access'
];
```

### Facebook API
```javascript
// Required permissions
const facebookScopes = [
  'pages_manage_posts',
  'pages_read_engagement',
  'pages_show_list'
];
```

### Instagram API
```javascript
// Required permissions
const instagramScopes = [
  'instagram_basic',
  'instagram_content_publish',
  'pages_show_list'
];
```

### LinkedIn API
```javascript
// Required permissions
const linkedinScopes = [
  'w_member_social',
  'r_liteprofile',
  'r_emailaddress'
];
```

## Security Best Practices

### Production Deployment

⚠️ **Important**: The current implementation stores credentials in localStorage for demonstration. In production:

1. **Backend API Service**
   ```javascript
   // Example Express.js OAuth handler
   app.get('/auth/:platform', (req, res) => {
     const { platform } = req.params;
     const authUrl = generateOAuthUrl(platform);
     res.redirect(authUrl);
   });

   app.get('/auth/:platform/callback', async (req, res) => {
     const { code } = req.query;
     const tokens = await exchangeCodeForTokens(platform, code);
     
     // Encrypt and store tokens securely
     await storeUserTokens(req.user.id, platform, tokens);
     res.redirect('/dashboard?connected=true');
   });
   ```

2. **Database Storage**
   ```sql
   CREATE TABLE user_social_accounts (
     id UUID PRIMARY KEY,
     user_id UUID REFERENCES users(id),
     platform VARCHAR(50) NOT NULL,
     encrypted_access_token TEXT,
     encrypted_refresh_token TEXT,
     expires_at TIMESTAMP,
     created_at TIMESTAMP DEFAULT NOW()
   );
   ```

3. **Environment Variables**
   ```bash
   # Social Media API Credentials
   TWITTER_CLIENT_ID=your_twitter_client_id
   TWITTER_CLIENT_SECRET=your_twitter_client_secret
   FACEBOOK_APP_ID=your_facebook_app_id
   FACEBOOK_APP_SECRET=your_facebook_app_secret
   INSTAGRAM_CLIENT_ID=your_instagram_client_id
   INSTAGRAM_CLIENT_SECRET=your_instagram_client_secret
   LINKEDIN_CLIENT_ID=your_linkedin_client_id
   LINKEDIN_CLIENT_SECRET=your_linkedin_client_secret

   # Encryption
   ENCRYPTION_KEY=your_32_character_encryption_key
   JWT_SECRET=your_jwt_secret_key
   ```

## Usage

### Account Management
1. **Connect Accounts**: Click "Manage Accounts" to connect social media platforms
2. **OAuth Flow**: Authenticate with each platform securely
3. **Permission Review**: Review and approve required permissions
4. **Multiple Accounts**: Connect multiple accounts per platform if needed

### Content Creation
1. **Select Content Type**: Choose between Image, Video, or Text-only posts
2. **Upload or Generate**: Upload existing media or use AI generation
3. **Add Text Overlay**: Customize text appearance and position (for media posts)
4. **Write Post Content**: Use the rich text editor with formatting tools
5. **Select Platforms**: Choose target social media platforms (only connected accounts)
6. **Schedule or Publish**: Post immediately or schedule for later

### AI Generation Options

#### Images
- **Styles**: Photorealistic, Artistic, Minimalist, Vintage, Futuristic
- **Sizes**: Square (1024×1024), Landscape (1792×1024), Portrait (1024×1792)
- **Quality**: Standard, HD
- **Models**: DALL-E 3, DALL-E 2

#### Videos
- **Duration**: 4-14 seconds
- **Resolution**: Up to 1280×768
- **Frame Rate**: 6-30 FPS
- **Motion Control**: Adjustable motion intensity

## API Rate Limits

### Platform Limits
- **Twitter**: 300 tweets per 15-minute window
- **Facebook**: 200 posts per hour per page
- **Instagram**: 25 posts per day per account
- **LinkedIn**: 100 posts per day per member

### AI Service Limits
- **OpenAI DALL-E**: 50 images per minute
- **Stability AI**: 150 images per minute
- **Replicate**: Varies by model
- **Runway ML**: Based on subscription tier

## Error Handling

The application includes comprehensive error handling for:
- **Network failures**: Automatic retry with exponential backoff
- **Token expiration**: Automatic refresh or re-authentication prompt
- **Rate limiting**: Queue management and user notifications
- **API errors**: User-friendly error messages and recovery suggestions

## Browser Compatibility

- Chrome 90+
- Firefox 88+
- Safari 14+
- Edge 90+

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

## License

MIT License - see LICENSE file for details.

## Support

For support and questions:
- Create an issue on GitHub
- Check the documentation
- Review the API integration guides

---

**Note**: This application handles sensitive user credentials. Always follow security best practices and comply with platform terms of service when implementing in production.