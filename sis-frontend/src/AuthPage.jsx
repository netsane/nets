import React, { useState } from 'react';
import { Box, TextField, Button, Typography, Paper } from '@mui/material';
import { useNavigate } from 'react-router-dom';
import { useAuth } from './AuthContext';

const AuthPage = () => {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [isRegistering, setIsRegistering] = useState(false);
  const [authMessage, setAuthMessage] = useState('');
  const navigate = useNavigate();
  const { login, register } = useAuth();

  const handleSubmit = async (event) => {
    event.preventDefault();
    setAuthMessage(''); // Clear previous messages immediately

    if (!username || !password) {
      setAuthMessage('Please enter both username and password.');
      return;
    }

    let result;
    if (isRegistering) {
      console.log('AuthPage: Calling register function with username:', username);
      result = await register(username, password);
    } else {
      console.log('AuthPage: Calling login function with username:', username);
      result = await login(username, password);
    }

    console.log('AuthPage: Login/Register attempt result:', result);
    console.log('AuthPage: Result success property:', result.success);

    if (result && result.success) {
      console.log('AuthPage: Login/Register successful, navigating to /dashboard');
      setAuthMessage(''); // Clear message on success before navigating
      navigate('/dashboard'); // Navigate to dashboard on successful login/registration
    } else {
      console.log('AuthPage: Login/Register failed or result.success is false. Displaying message:', result?.message);
      setAuthMessage(result?.message || 'An unknown error occurred during authentication.');
    }
  };

  return (
    <Box
      sx={{
        height: '100vh',
        width: '100vw',
        display: 'flex',
        justifyContent: 'center',
        alignItems: 'center',
        backgroundImage: 'url(/images/campus_photo.jpg)', // Background image for the auth page
        backgroundSize: 'cover',
        backgroundPosition: 'center',
        filter: 'brightness(0.7) contrast(0.7)', // Apply filter to background image
        transition: 'filter 0.5s ease-in-out', // Smooth transition for filter changes
      }}
    >
      <Paper
        elevation={0} // No shadow for the paper itself, relying on image shadow
        sx={{
          p: 4,
          borderRadius: '12px',
          maxWidth: '400px',
          width: '100%',
          bgcolor: 'transparent', // Transparent background for the form container
          boxShadow: 'none', // No shadow on the paper itself
          display: 'flex',
          flexDirection: 'column',
          gap: 2,
          alignItems: 'center',
        }}
      >
        <img
          src="/images/logo.jpg" // University logo
          alt="Unity University Logo"
          style={{
            width: 100,
            height: 100,
            borderRadius: '50%',
            marginBottom: '16px',
            boxShadow: '0 0 15px 5px rgba(255, 255, 255, 0.5)', // White glow around the logo
            objectFit: 'cover',
          }}
          onError={(e) => { // Fallback for image loading errors
            console.error("Failed to load logo from:", e.target.src);
            e.target.src = 'https://placehold.co/100x100/cccccc/000000?text=Logo';
            e.target.alt = 'Logo missing';
          }}
        />

        <Typography
          variant="h5"
          component="h1"
          gutterBottom
          sx={{
            fontWeight: 'bold',
            color: 'rgba(0, 0, 0, 0.9)', // Dark text for contrast
            textShadow: '0 0 5px rgba(255, 255, 255, 0.8), 0 0 10px rgba(255, 255, 255, 0.6)', // White glow for text
          }}
        >
          Integrated SIS
        </Typography>
        <Typography
          variant="body2"
          sx={{
            mb: 2,
            textAlign: 'center',
            color: 'rgba(0, 0, 0, 0.8)',
            textShadow: '0 0 3px rgba(255, 255, 255, 0.7), 0 0 6px rgba(255, 255, 255, 0.5)',
          }}
        >
          Welcome to the SIS Dashboard! Unity University, established in 1991 in Addis Ababa, Ethiopia, is a pioneer private higher education institution. It was the first privately owned college in Ethiopia to be awarded full university status and offer postgraduate programs. The university is committed to providing high-quality education, fostering academic excellence, and contributing to national development through diverse programs in fields like Business Administration, Information Technology, Law, Public Health, and Engineering. Please log in or register to access the content.
        </Typography>
        {authMessage && (
          <Typography color="error" sx={{ mb: 1, textAlign: 'center' }}>
            {authMessage}
          </Typography>
        )}
        <TextField
          placeholder="Username"
          variant="outlined"
          fullWidth
          value={username}
          onChange={(e) => setUsername(e.target.value)}
          sx={{
            '& .MuiInputBase-root': {
              borderRadius: '8px',
              bgcolor: 'rgba(255, 255, 255, 0.1)', // Transparent input background
              color: 'rgba(0, 0, 0, 0.9)', // Dark input text
              textShadow: '0 0 2px rgba(255, 255, 255, 0.8)', // White glow for input text
            },
            '& .MuiOutlinedInput-root': {
              '& fieldset': { borderColor: 'rgba(25, 118, 210, 0.5)' }, // Blue border
              '&:hover fieldset': { borderColor: 'rgba(25, 118, 210, 0.8)' }, // Darker blue on hover
              '&.Mui-focused fieldset': { borderColor: 'primary.main' }, // Primary color when focused
            },
          }}
        />
        <TextField
          placeholder="Password"
          variant="outlined"
          type="password"
          fullWidth
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          sx={{
            '& .MuiInputBase-root': {
              borderRadius: '8px',
              bgcolor: 'rgba(255, 255, 255, 0.1)',
              color: 'rgba(0, 0, 0, 0.9)',
              textShadow: '0 0 2px rgba(255, 255, 255, 0.8)',
            },
            '& .MuiOutlinedInput-root': {
              '& fieldset': { borderColor: 'rgba(25, 118, 210, 0.5)' },
              '&:hover fieldset': { borderColor: 'rgba(25, 118, 210, 0.8)' },
              '&.Mui-focused fieldset': { borderColor: 'primary.main' },
            },
          }}
        />
        <Button
          variant="outlined" // Outlined button for a lighter look
          color="primary"
          fullWidth
          onClick={handleSubmit}
          sx={{
            py: 1.5,
            fontWeight: 'bold',
            borderRadius: '8px',
            bgcolor: 'transparent', // Transparent background
            color: 'primary.main', // Blue text
            border: '1px solid',
            borderColor: 'primary.main',
            transition: 'background-color 0.3s ease, box-shadow 0.3s ease, border-color 0.3s ease, color 0.3s ease',
            '&:hover': {
              bgcolor: 'rgba(25, 118, 210, 0.1)', // Slight blue tint on hover
              boxShadow: '0 0 15px 3px rgba(25, 118, 210, 0.7)', // Blue glow on hover
              borderColor: 'primary.dark',
              color: 'primary.dark',
            },
            '&:active': {
              bgcolor: 'rgba(25, 118, 210, 0.2)', // More tint on active
              boxShadow: '0 0 20px 5px rgba(25, 118, 210, 0.9)', // More intense blue glow on active
              borderColor: 'primary.dark',
              color: 'primary.dark',
            },
            '&.Mui-disabled': {
              bgcolor: 'rgba(0, 0, 0, 0.12)',
              color: 'rgba(0, 0, 0, 0.26)',
              borderColor: 'rgba(0, 0, 0, 0.12)',
              boxShadow: 'none',
            }
          }}
        >
          {isRegistering ? 'Register' : 'Login'}
        </Button>
        <Button
          variant="text"
          color="secondary"
          fullWidth
          onClick={() => {
            setIsRegistering((prev) => !prev);
            setAuthMessage(''); // Clear message when switching form
            setUsername(''); // Clear input fields
            setPassword('');
          }}
          sx={{
            mt: 1,
            textTransform: 'none', // Keep text as is
            color: 'secondary.main', // Secondary color (e.g., purple)
            textShadow: '0 0 5px rgba(255, 255, 255, 0.7), 0 0 10px rgba(255, 255, 255, 0.5)', // White glow
            transition: 'color 0.3s ease, text-shadow 0.3s ease',
            '&:hover': {
              color: 'secondary.dark',
              textShadow: '0 0 8px rgba(255, 255, 255, 0.9), 0 0 15px rgba(255, 255, 255, 0.7)',
              bgcolor: 'transparent' // Keep transparent on hover
            },
            '&:active': {
              color: 'secondary.dark',
              textShadow: '0 0 10px rgba(255, 255, 255, 1), 0 0 20px rgba(255, 255, 255, 0.8)',
              bgcolor: 'transparent' // Keep transparent on active
            }
          }}
        >
          {isRegistering ? 'Already have an account? Login' : 'Need an account? Register'}
        </Button>
      </Paper>
    </Box>
  );
};

export default AuthPage;
