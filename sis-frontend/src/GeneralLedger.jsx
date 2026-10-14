import React, { useState, useEffect } from 'react';
import {
  Box, Typography, Paper, Button, MenuItem,
  Table, TableBody, TableCell, TableContainer, TableHead, TableRow,
  CircularProgress, Alert, AppBar, Toolbar, IconButton, useTheme, TextField
} from '@mui/material';
import { useNavigate } from 'react-router-dom';
import { useAuth } from './AuthContext';
import axios from 'axios';
import ArrowBackIcon from '@mui/icons-material/ArrowBack';
import AddCircleOutlineIcon from '@mui/icons-material/AddCircleOutline'; // Re-using this for "Generate Report"
import dayjs from 'dayjs';
import { DatePicker, LocalizationProvider } from '@mui/x-date-pickers';
import { AdapterDayjs } from '@mui/x-date-pickers/AdapterDayjs';

const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || 'http://localhost:3001/api';

function GeneralLedger() {
  const navigate = useNavigate();
  const { logout, currentUser } = useAuth();
  const theme = useTheme();

  // State for report parameters
  const [selectedAccountId, setSelectedAccountId] = useState('');
  const [startDate, setStartDate] = useState(dayjs().startOf('year')); // Default to start of current year
  const [endDate, setEndDate] = useState(dayjs()); // Default to current date

  // State for data and UI feedback
  const [accounts, setAccounts] = useState([]); // To populate the account dropdown
  const [generalLedgerData, setGeneralLedgerData] = useState(null);
  const [isGeneratingReport, setIsGeneratingReport] = useState(false);
  const [reportError, setReportError] = useState('');
  const [reportSuccessMessage, setReportSuccessMessage] = useState('');

  console.log('GeneralLedger.jsx: Component rendering...');

  // Fetch Accounts for the dropdown
  const fetchAccounts = async () => {
    try {
      const token = localStorage.getItem('token');
      if (!token) {
        console.error('GeneralLedger.jsx: No authentication token found.');
        setReportError('Authentication token missing. Please log in again.');
        logout();
        return;
      }
      const config = { headers: { 'Authorization': `Bearer ${token}` } };
      const response = await axios.get(`${API_BASE_URL}/accounts`, config);
      setAccounts(response.data);
      if (response.data.length > 0 && !selectedAccountId) {
        setSelectedAccountId(response.data[0].id); // Select the first account by default
      }
    } catch (err) {
      console.error('GeneralLedger.jsx: Error fetching accounts:', err.response?.data || err.message);
      setReportError(err.response?.data?.message || 'Failed to load accounts.');
      if (err.response && (err.response.status === 401 || err.response.status === 403)) {
        logout();
      }
    }
  };

  // Fetch General Ledger Data
  const fetchGeneralLedger = async () => {
    setIsGeneratingReport(true);
    setReportError('');
    setReportSuccessMessage('');
    setGeneralLedgerData(null); // Clear previous report

    if (!selectedAccountId || !startDate || !endDate) {
      setReportError('Please select an Account, Start Date, and End Date.');
      setIsGeneratingReport(false);
      return;
    }

    console.log(`GeneralLedger.jsx: Fetching General Ledger for Account ${selectedAccountId} from ${startDate.format('YYYY-MM-DD')} to ${endDate.format('YYYY-MM-DD')}`);

    try {
      const token = localStorage.getItem('token');
      if (!token) {
        setReportError('Authentication token missing. Please log in again.');
        logout();
        return;
      }

      const config = {
        headers: { 'Authorization': `Bearer ${token}` },
        params: {
          accountId: selectedAccountId,
          startDate: startDate.format('YYYY-MM-DD'),
          endDate: endDate.format('YYYY-MM-DD')
        }
      };

      console.log(`GeneralLedger.jsx: Making GET request to ${API_BASE_URL}/transactions/reports/general-ledger with params:`, config.params);
      const response = await axios.get(`${API_BASE_URL}/transactions/reports/general-ledger`, config);

      setGeneralLedgerData(response.data);
      setReportSuccessMessage('General Ledger generated successfully!');
      console.log('GeneralLedger.jsx: General Ledger data received:', response.data);

    } catch (err) {
      console.error('GeneralLedger.jsx: Error generating General Ledger:', err.response?.data || err.message);
      setReportError(err.response?.data?.message || 'Failed to generate General Ledger.');
      if (err.response && (err.response.status === 401 || err.response.status === 403)) {
        logout();
      }
    } finally {
      setIsGeneratingReport(false);
      console.log('GeneralLedger.jsx: Report generation complete. Loading state:', false);
    }
  };

  // Initial data fetch on component mount
  useEffect(() => {
    console.log('GeneralLedger.jsx: useEffect - Initial data fetch for accounts.');
    fetchAccounts();
  }, [logout]);

  // Fetch report when parameters change or accounts are loaded
  useEffect(() => {
    if (selectedAccountId && accounts.length > 0) {
      console.log('GeneralLedger.jsx: useEffect - Fetching report due to param change or accounts loaded.');
      fetchGeneralLedger();
    }
  }, [selectedAccountId, startDate, endDate, accounts.length, logout]); // Depend on accounts.length to ensure dropdown is populated

  const selectedAccountName = accounts.find(acc => acc.id === selectedAccountId)?.name || 'N/A';
  const selectedAccountType = accounts.find(acc => acc.id === selectedAccountId)?.type || 'N/A';

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
      <AppBar position="static" sx={{
        bgcolor: 'rgba(0,0,0,0.6)',
        boxShadow: 'none',
        flexShrink: 0
      }}>
        <Toolbar>
          <IconButton edge="start" color="inherit" aria-label="back" onClick={() => navigate('/financial-dashboard')}>
            <ArrowBackIcon />
          </IconButton>
          <Typography variant="h6" component="div" sx={{ flexGrow: 1, color: 'white', textShadow: '0 0 5px rgba(255,255,255,0.8)' }}>
            Financial Reports
          </Typography>
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 2 }}>
            {/* Navigation to other reports */}
            <Button
              variant="outlined"
              color="inherit"
              size="small"
              onClick={() => navigate('/income-statement')}
              sx={{ textTransform: 'none', fontWeight: 'bold', borderRadius: '8px', color: 'white', borderColor: 'white' }}
            >
              Income Statement
            </Button>
            <Button
              variant="outlined"
              color="inherit"
              size="small"
              onClick={() => navigate('/balance-sheet')}
              sx={{ textTransform: 'none', fontWeight: 'bold', borderRadius: '8px', color: 'white', borderColor: 'white' }}
            >
              Balance Sheet
            </Button>
            <Button
              variant="outlined"
              color="inherit"
              size="small"
              onClick={() => navigate('/cash-flow-statement')}
              sx={{ textTransform: 'none', fontWeight: 'bold', borderRadius: '8px', color: 'white', borderColor: 'white' }}
            >
              Cash Flow
            </Button>
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
            Generate General Ledger for:
          </Typography>

          {/* Account Select */}
          <TextField
            select
            label="Select Account"
            value={selectedAccountId}
            onChange={(e) => setSelectedAccountId(e.target.value)}
            size="small"
            variant="outlined"
            sx={{ width: '220px', '& .MuiInputBase-root': { borderRadius: '8px' } }}
          >
            {accounts.map((account) => (
              <MenuItem key={account.id} value={account.id}>
                {account.name} ({account.type})
              </MenuItem>
            ))}
          </TextField>

          {/* Date Range Pickers */}
          <LocalizationProvider dateAdapter={AdapterDayjs}>
            <DatePicker
              label="Start Date"
              value={startDate}
              onChange={(newValue) => setStartDate(newValue)}
              slotProps={{
                textField: {
                  size: 'small',
                  variant: 'outlined',
                  sx: { width: '180px', '& .MuiInputBase-root': { borderRadius: '8px' } }
                }
              }}
            />
            <DatePicker
              label="End Date"
              value={endDate}
              onChange={(newValue) => setEndDate(newValue)}
              slotProps={{
                textField: {
                  size: 'small',
                  variant: 'outlined',
                  sx: { width: '180px', '& .MuiInputBase-root': { borderRadius: '8px' } }
                }
              }}
            />
          </LocalizationProvider>

          <Button
            variant="contained"
            color="primary"
            onClick={fetchGeneralLedger}
            disabled={isGeneratingReport}
            startIcon={isGeneratingReport ? <CircularProgress size={18} color="inherit" /> : <AddCircleOutlineIcon />}
            sx={{ py: 1.2, fontWeight: 'bold', borderRadius: '8px', textTransform: 'none' }}
          >
            {isGeneratingReport ? 'Generating...' : 'Generate Report'}
          </Button>
        </Paper>

        {/* Error and Success Messages */}
        {reportError && <Alert severity="error" sx={{ mb: 2 }}>{reportError}</Alert>}
        {reportSuccessMessage && <Alert severity="success" sx={{ mb: 2 }}>{reportSuccessMessage}</Alert>}

        {/* General Ledger Display */}
        {generalLedgerData && (
          <Paper elevation={3} sx={{
            p: 3,
            mb: 3,
            borderRadius: '12px',
            bgcolor: 'rgba(255,255,255,0.9)',
            boxShadow: '0 4px 15px rgba(0,0,0,0.1)',
            width: '100%'
          }}>
            <Typography variant="h6" gutterBottom sx={{ fontWeight: 'bold', color: theme.palette.text.primary, textAlign: 'center' }}>
              General Ledger for {selectedAccountName} ({selectedAccountType})
            </Typography>
            <Typography variant="subtitle1" gutterBottom sx={{ textAlign: 'center', mb: 2 }}>
              From {startDate.format('MM/DD/YYYY')} to {endDate.format('MM/DD/YYYY')}
            </Typography>

            <TableContainer sx={{ maxHeight: 600 }}>
              <Table size="small" stickyHeader>
                <TableHead sx={{ bgcolor: theme.palette.primary.light }}>
                  <TableRow>
                    <TableCell sx={{ fontWeight: 'bold', color: theme.palette.primary.contrastText }}>Date</TableCell>
                    <TableCell sx={{ fontWeight: 'bold', color: theme.palette.primary.contrastText }}>Description</TableCell>
                    <TableCell sx={{ fontWeight: 'bold', color: theme.palette.primary.contrastText }}>Debit</TableCell>
                    <TableCell sx={{ fontWeight: 'bold', color: theme.palette.primary.contrastText }}>Credit</TableCell>
                    <TableCell sx={{ fontWeight: 'bold', color: theme.palette.primary.contrastText }}>Balance</TableCell>
                  </TableRow>
                </TableHead>
                <TableBody>
                  <TableRow sx={{ bgcolor: theme.palette.grey[100] }}>
                    <TableCell colSpan={4} sx={{ fontWeight: 'bold' }}>Beginning Balance</TableCell>
                    <TableCell align="right" sx={{ fontWeight: 'bold' }}>${parseFloat(generalLedgerData.beginningBalance).toFixed(2)}</TableCell>
                  </TableRow>
                  {generalLedgerData.transactions.length > 0 ? (
                    generalLedgerData.transactions.map((transaction, index) => (
                      <TableRow key={transaction.id} sx={{ bgcolor: index % 2 ? 'rgba(0,0,0,0.02)' : 'inherit' }}>
                        <TableCell>{dayjs(transaction.transaction_date).format('MM/DD/YYYY')}</TableCell>
                        <TableCell>{transaction.description}</TableCell>
                        <TableCell align="right">{transaction.type === 'debit' ? `$${parseFloat(transaction.amount).toFixed(2)}` : '-'}</TableCell>
                        <TableCell align="right">{transaction.type === 'credit' ? `$${parseFloat(transaction.amount).toFixed(2)}` : '-'}</TableCell>
                        <TableCell align="right">${parseFloat(transaction.runningBalance).toFixed(2)}</TableCell>
                      </TableRow>
                    ))
                  ) : (
                    <TableRow>
                      <TableCell colSpan={5} sx={{ textAlign: 'center', py: 3 }}>No transactions found for this period.</TableCell>
                    </TableRow>
                  )}
                  <TableRow sx={{ bgcolor: theme.palette.primary.light, '& > td': { fontWeight: 'bold', color: theme.palette.primary.contrastText } }}>
                    <TableCell colSpan={4}>Ending Balance</TableCell>
                    <TableCell align="right">${parseFloat(generalLedgerData.endingBalance).toFixed(2)}</TableCell>
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

export default GeneralLedger;
