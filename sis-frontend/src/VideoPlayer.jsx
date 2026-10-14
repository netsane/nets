import React from 'react';
import { Box, Typography } from '@mui/material';

// This is a minimal, self-contained video player component
const VideoPlayer = ({ videoUrl, videoTitle, posterImage }) => {
  return (
    <Box sx={{
      // Container for the video and title
      bgcolor: 'rgba(0, 0, 0, 0.5)',
      borderRadius: '12px',
      overflow: 'hidden', // Ensures video corners are rounded
      boxShadow: '0 4px 8px rgba(0,0,0,0.3)',
      width: '100%',
      height: '100%',
      display: 'flex',
      flexDirection: 'column',
      position: 'relative'
    }}>
      <Typography 
        variant="h6" 
        sx={{ 
          p: 2, 
          color: 'white', 
          fontWeight: 'bold', 
          textAlign: 'center',
          backgroundColor: 'rgba(0, 0, 0, 0.4)',
        }}
      >
        {videoTitle}
      </Typography>
      <video
        width="100%"
        height="100%"
        controls // This attribute is key to enabling the native player controls
        poster={posterImage}
        style={{ display: 'block', flexGrow: 1 }}
      >
        <source src={videoUrl} type="video/mp4" />
        Your browser does not support the video tag.
      </video>
    </Box>
  );
};

export default VideoPlayer;
