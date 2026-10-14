import React, { useState, useEffect } from 'react';
import {
  Box, Typography, Paper, TextField, Button, MenuItem,
  Table, TableBody, TableCell, TableContainer, TableHead, TableRow,
  CircularProgress, Alert, AppBar, Toolbar, IconButton, useTheme, Select
} from '@mui/material';
import { useNavigate } from 'react-router-dom'; // <--- ADD THIS IMPORT
import { useAuth } from './AuthContext';
import axios from 'axios';
import ArrowBackIcon from '@mui/icons-material/ArrowBack';
import AddCircleOutlineIcon from '@mui/icons-material/AddCircleOutline';
import dayjs from 'dayjs';
import { DatePicker, LocalizationProvider } from '@mui/x-date-pickers';
import { AdapterDayjs } from '@mui/x-date-pickers/AdapterDayjs';

const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || 'http://localhost:3001/api';

function BalanceSheet() {
  const navigate = useNavigate(); // This line now has access to the hook
  const { logout, currentUser } = useAuth();
  const theme = useTheme();

  // --- Reporting State ---
  // State for the Balance Sheet date selection
  const [reportDate, setReportDate] = useState(dayjs(new Date())); // State for the 'as of' date

  const [balanceSheet, setBalanceSheet] = useState(null);
  const [isGeneratingReport, setIsGeneratingReport] = useState(false);
  const [reportError, setReportError] = useState('');
  const [reportSuccessMessage, setReportSuccessMessage] = useState('');

  // Fetch Balance Sheet Data
  const fetchBalanceSheet = async () => {
    setIsGeneratingReport(true);
    setReportError('');
    setReportSuccessMessage('');
    setBalanceSheet(null);

    if (!reportDate) {
      setReportError('Please select a date to generate the Balance Sheet.');
      console.warn('BalanceSheet.jsx: Cannot generate report, date is null.');
      setIsGeneratingReport(false);
      return;
    }

    // Extract year and month from reportDate for API call
    const reportYearNum = reportDate.year(); // This is not used in the params for balance sheet, but kept for consistency
    const reportMonthNum = reportDate.month() + 1; // This is not used in the params for balance sheet, but kept for consistency

    console.log(`BalanceSheet.jsx: Fetching Balance Sheet for Date: ${reportDate.format('YYYY-MM-DD')}`);

    try {
      const token = localStorage.getItem('token');
      if (!token) {
        console.error('BalanceSheet.jsx: No authentication token found.');
        setReportError('Authentication token missing. Please log in again.');
        logout();
        return;
      }

      const config = {
        headers: { 'Authorization': `Bearer ${token}` },
        // Pass the full date, backend will extract month/year or just use it for filtering
        params: { date: reportDate.format('YYYY-MM-DD') }
      };

      console.log(`BalanceSheet.jsx: Making GET request to ${API_BASE_URL}/transactions/reports/balance-sheet with params:`, config.params);
      const response = await axios.get(`${API_BASE_URL}/transactions/reports/balance-sheet`, config);
      setBalanceSheet(response.data);
      setReportSuccessMessage('Balance Sheet generated successfully!');
      console.log('BalanceSheet.jsx: Balance Sheet received:', response.data);

    } catch (err) {
      console.error('BalanceSheet.jsx: Error generating Balance Sheet:', err.response?.data || err.message);
      setReportError(err.response?.data?.message || 'Failed to generate Balance Sheet.');
      if (err.response && (err.response.status === 401 || err.response.status === 403)) {
        console.log('BalanceSheet.jsx: Unauthorized error generating report, logging out.');
        logout();
      }
    } finally {
      setIsGeneratingReport(false);
      console.log('BalanceSheet.jsx: Report generation complete. Loading state:', false);
    }
  };

  // Corrected: useEffect to use reportDate for fetching
  useEffect(() => {
    console.log('BalanceSheet.jsx: useEffect triggered. Fetching report for date:', reportDate.format('YYYY-MM-DD'));
    fetchBalanceSheet();
  }, [reportDate, logout]); // Depend on reportDate to re-fetch when date changes

  // Handler for date changes in the picker
  const handleDateChange = (newValue) => {
    if (newValue) {
      setReportDate(newValue); // Update the reportDate state with the new dayjs object
      console.log(`BalanceSheet.jsx: Date changed to: ${newValue.format('YYYY-MM-DD')}`);
    }
  };


  return (
    <Box sx={{
      height: '100vh',
      width: '100vw',
      display: 'flex',
      flexDirection: 'column',
      bgcolor: theme.palette.grey[100],
      backgroundImage: 'url(/images/campus_photo.jpg)',
      backgroundSize: 'cover',
      backgroundPosition: 'center',
      filter: 'brightness(0.6) contrast(0.8)',
      transition: 'filter 0.5s ease-in-out',
    }}>
      {/* AppBar for navigation and logout */}
      <AppBar position="static" sx={{ bgcolor: 'rgba(0,0,0,0.6)', boxShadow: 'none', flexShrink: 0 }}>
        <Toolbar>
          <IconButton edge="start" color="inherit" aria-label="back" onClick={() => navigate('/financial-dashboard')}>
            <ArrowBackIcon />
          </IconButton>
          <Typography variant="h6" component="div" sx={{ flexGrow: 1, color: 'white', textShadow: '0 0 5px rgba(255,255,255,0.8)' }}>
            Financial Reports
          </Typography>
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 2 }}>
            {/* --- NAVIGATION TO REPORTS --- */}
            <Button
              variant="outlined"
              color="inherit"
              size="small"
              onClick={() => navigate('/income-statement')} // Navigate to Income Statement
              sx={{ textTransform: 'none', fontWeight: 'bold', borderRadius: '8px', color: 'white', borderColor: 'white' }}
            >
              Income Statement
            </Button>
            {/* --- END NAVIGATION --- */}
            {currentUser && (
              <Button
                onClick={logout}
                variant="outlined"
                color="inherit"
                size="small"
                sx={{ ml: 2, textTransform: 'none', fontWeight: 'bold', borderRadius: '8px', color: 'white', borderColor: 'white' }}
              >
                Logout ({currentUser?.username})
              </Button>
            )}
          </Box>
        </Toolbar>
      </AppBar>

      {/* Main content area, scrollable */}
      <Box sx={{
        p: 3,
        flexGrow: 1,
        overflowY: 'scroll',
        width: '100%',
        minHeight: 0,
        '&::-webkit-scrollbar': { width: '8px' },
        '&::-webkit-scrollbar-track': { background: 'rgba(255,255,255,0.2)', borderRadius: '4px' },
        '&::-webkit-scrollbar-thumb': { background: 'rgba(0,0,0,0.6)', borderRadius: '4px' },
        '&::-webkit-scrollbar-thumb:hover': { background: 'rgba(0,0,0,0.8)' },
        pb: 5,
      }}>
        {/* Report Generation Controls */}
        <Paper elevation={3} sx={{
          p: 3,
          mb: 3,
          borderRadius: '12px',
          bgcolor: 'rgba(255,255,255,0.9)',
          boxShadow: '0 4px 15px rgba(0,0,0,0.1)',
          display: 'flex',
          alignItems: 'center',
          gap: 2,
          flexWrap: 'wrap',
          justifyContent: 'center'
        }}>
          <Typography variant="h6" sx={{ fontWeight: 'bold', color: theme.palette.text.primary }}>
            Generate Balance Sheet as of:
          </Typography>

          {/* Date Selector (for Balance Sheet) */}
          <LocalizationProvider dateAdapter={AdapterDayjs}>
            <DatePicker
              label="Select As Of Date"
              views={['year', 'month', 'day']}
              value={reportDate} // Use the reportDate state
              onChange={handleDateChange} // Use the new handler
              slotProps={{
                textField: {
                  size: 'small',
                  variant: 'outlined',
                  sx: { width: '220px', '& .MuiInputBase-root': { borderRadius: '8px' } }
                }
              }}
            />
          </LocalizationProvider>

          <Button
            variant="contained"
            color="primary"
            onClick={fetchBalanceSheet}
            disabled={isGeneratingReport}
            startIcon={isGeneratingReport ? <CircularProgress size={18} color="inherit" /> : <AddCircleOutlineIcon />}
            sx={{ py: 1.2, fontWeight: 'bold', borderRadius: '8px', textTransform: 'none' }}
          >
            {isGeneratingReport ? 'Generating...' : 'Generate Balance Sheet'}
          </Button>
        </Paper>

        {/* Error and Success Messages */}
        {reportError && <Alert severity="error" sx={{ mb: 2 }}>{reportError}</Alert>}
        {reportSuccessMessage && <Alert severity="success" sx={{ mb: 2 }}>{reportSuccessMessage}</Alert>}


        {/* Balance Sheet Display */}
        {balanceSheet && ( // Only render if balanceSheet data exists
          <Paper elevation={3} sx={{
            p: 3,
            mb: 3,
            borderRadius: '12px',
            bgcolor: 'rgba(255,255,255,0.9)',
            boxShadow: '0 4px 15px rgba(0,0,0,0.1)',
            width: '100%'
          }}>
            <Typography variant="h6" gutterBottom sx={{ fontWeight: 'bold', color: theme.palette.text.primary, textAlign: 'center' }}>
              Balance Sheet as of {new Date(balanceSheet.date).toLocaleDateString()}
            </Typography>
            <TableContainer sx={{ maxHeight: 600 }}>
              <Table size="small" stickyHeader>
                <TableHead sx={{ bgcolor: theme.palette.primary.light }}>
                  <TableRow>
                    <TableCell sx={{ fontWeight: 'bold', color: theme.palette.primary.contrastText }}>Account</TableCell>
                    <TableCell sx={{ fontWeight: 'bold', color: theme.palette.primary.contrastText }}>Amount</TableCell>
                  </TableRow>
                </TableHead>
                <TableBody>
                  {/* Assets Section */}
                  <TableRow>
                    <TableCell colSpan={2} sx={{ bgcolor: theme.palette.grey[200], fontWeight: 'bold', borderBottom: '1px solid #ccc' }}>
                      Assets
                    </TableCell>
                  </TableRow>
                  {Object.entries(balanceSheet.assets?.assetItems || {}).map(([accountName, amount]) => (
                    <TableRow key={accountName}>
                      <TableCell>{accountName}</TableCell>
                      <TableCell align="right">${parseFloat(amount).toFixed(2)}</TableCell>
                    </TableRow>
                  ))}
                  <TableRow>
                    <TableCell sx={{ fontWeight: 'bold' }}>Total Assets</TableCell>
                    <TableCell align="right" sx={{ fontWeight: 'bold' }}>${parseFloat(balanceSheet.assets?.totalAssets).toFixed(2)}</TableCell>
                  </TableRow>

                  <TableRow>
                    <TableCell colSpan={2} sx={{ height: '20px' }}></TableCell> {/* Spacer row */}
                  </TableRow>

                  {/* Liabilities Section */}
                  <TableRow>
                    <TableCell colSpan={2} sx={{ bgcolor: theme.palette.grey[200], fontWeight: 'bold', borderBottom: '1px solid #ccc' }}>
                      Liabilities
                    </TableCell>
                  </TableRow>
                  {Object.entries(balanceSheet.liabilities?.liabilityItems || {}).map(([accountName, amount]) => (
                    <TableRow key={accountName}>
                      <TableCell>{accountName}</TableCell>
                      <TableCell align="right">${parseFloat(amount).toFixed(2)}</TableCell>
                    </TableRow>
                  ))}
                  <TableRow>
                    <TableCell sx={{ fontWeight: 'bold' }}>Total Liabilities</TableCell>
                    <TableCell align="right" sx={{ fontWeight: 'bold' }}>${parseFloat(balanceSheet.liabilities?.totalLiabilities).toFixed(2)}</TableCell>
                  </TableRow>

                  <TableRow>
                    <TableCell colSpan={2} sx={{ height: '20px' }}></TableCell> {/* Spacer row */}
                  </TableRow>

                  {/* Equity Section */}
                  <TableRow>
                    <TableCell colSpan={2} sx={{ bgcolor: theme.palette.grey[200], fontWeight: 'bold', borderBottom: '1px solid #ccc' }}>
                      Equity
                    </TableCell>
                  </TableRow>
                  {Object.entries(balanceSheet.equity?.equityItems || {}).map(([accountName, amount]) => (
                    <TableRow key={accountName}>
                      <TableCell>{accountName}</TableCell>
                      <TableCell align="right">${parseFloat(amount).toFixed(2)}</TableCell>
                    </TableRow>
                  ))}
                  <TableRow>
                    <TableCell sx={{ fontWeight: 'bold' }}>Total Equity</TableCell>
                    <TableCell align="right" sx={{ fontWeight: 'bold' }}>${parseFloat(balanceSheet.equity?.totalEquity).toFixed(2)}</TableCell>
                  </TableRow>

                  <TableRow>
                    <TableCell colSpan={2} sx={{ height: '20px' }}></TableCell> {/* Spacer row */}
                  </TableRow>

                  {/* Total Liabilities and Equity Section */}
                  <TableRow sx={{ bgcolor: (Math.abs(balanceSheet.assets?.totalAssets - balanceSheet.totalLiabilitiesAndEquity) < 0.01) ? '#e8f5e9' : '#ffebee', '& > td': { fontWeight: 'bold' } }}>
                    <TableCell>Total Liabilities & Equity</TableCell>
                    <TableCell align="right">${parseFloat(balanceSheet.totalLiabilitiesAndEquity).toFixed(2)}</TableCell>
                  </TableRow>

                  {/* Final Check: Total Assets vs. Total Liabilities & Equity */}
                  <TableRow sx={{ bgcolor: (Math.abs(balanceSheet.assets?.totalAssets - balanceSheet.totalLiabilitiesAndEquity) < 0.01) ? '#c8e6c9' : '#ffebee', '& > td': { fontWeight: 'bold' } }}>
                    <TableCell>Total Assets</TableCell>
                    <TableCell align="right">${parseFloat(balanceSheet.assets?.totalAssets).toFixed(2)}</TableCell>
                  </TableRow>
                </TableBody>
              </Table>
            </TableContainer>
          </Paper>
        )}
      </Box>
    </Box>
  );
}

export default BalanceSheet;
