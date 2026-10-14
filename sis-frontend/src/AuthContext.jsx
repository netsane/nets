import React, { createContext, useContext, useState, useEffect } from 'react';
import axios from 'axios';

const AuthContext = createContext();

const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || 'http://localhost:3001/api';

export const AuthProvider = ({ children }) => {
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [currentUser, setCurrentUser] = useState(null);
  const [loading, setLoading] = useState(true); // To manage initial auth check loading state

  useEffect(() => {
    console.log('AuthContext: Initializing, checking for token and user...');
    const token = localStorage.getItem('token');
    const user = localStorage.getItem('user');

    if (token && user) {
      console.log('AuthContext: localStorage token: exists');
      console.log('AuthContext: localStorage user: exists');
      try {
        const parsedUser = JSON.parse(user);
        setCurrentUser(parsedUser);
        setIsLoggedIn(true); // This sets the state
        console.log('AuthContext: isLoggedIn set to TRUE after parsing user.'); // Added log
        // Set default Authorization header for all Axios requests
        axios.defaults.headers.common['Authorization'] = `Bearer ${token}`;
        console.log('AuthContext: User found in localStorage, setting headers. Parsed user:', parsedUser);
      } catch (e) {
        console.error('AuthContext: Failed to parse user from localStorage:', e);
        logout(); // Clear invalid data
      }
    } else {
      console.log('AuthContext: No token or user found in localStorage.');
      setIsLoggedIn(false);
      setCurrentUser(null);
      delete axios.defaults.headers.common['Authorization']; // Clear header if no token
    }
    setLoading(false); // Initial check complete
    // This log will now show the state from the *previous* render,
    // but the separate useEffect below will show the *updated* state.
    console.log('AuthContext: Initial auth check completed. (State might not be fully updated yet)');
  }, []); // Empty dependency array means this runs once on mount

  // This useEffect will fire whenever isLoggedIn state actually changes
  useEffect(() => {
    console.log('AuthContext: isLoggedIn state *actually* changed to:', isLoggedIn);
    if (!loading) { // Only log routing decision after initial loading is done
      if (isLoggedIn) {
        console.log('AuthContext: User IS logged in. Should display dashboard.');
      } else {
        console.log('AuthContext: User is NOT logged in. Should show login page.');
      }
    }
  }, [isLoggedIn, loading]); // Depend on isLoggedIn and loading
// Inside AuthProvider component, update the 'login' function

const login = async (username, password) => {
    try {
        // ... (Mocking a successful login response for this environment)
        const token = 'mock-jwt-token-12345';
        let role = DEFAULT_ROLE;

        // --- NEW MOCK ROLE MAPPING ---
        if (username === 'admin') role = 'admin';
        else if (username === 'hr_manager') role = 'hr_manager';
        else if (username === 'hr_staff') role = 'hr_staff';
        else if (username === 'recruiter') role = 'recruiter';
        else if (username === 'hr_data_entry') role = 'hr_data_entry';
        else if (username === 'comp_assessor') role = 'comp_assessor';
        else if (username === 'appraisal_admin') role = 'appraisal_admin';
        else if (username === 'leave_processor') role = 'leave_processor';
        else if (username === 'chief_discipline') role = 'disciplinary_officer';
        else if (username === 'awards_admin') role = 'awards_admin';
        else if (username === 'transfer_officer') role = 'transfer_officer';
        else if (username === 'training_officer') role = 'training_officer';
        else if (username === 'insurance_admin') role = 'insurance_admin';
        else if (username === 'report_viewer') role = 'report_viewer';
        // --- END NEW MOCK ROLE MAPPING ---
        
        const user = { username: username, role: role };

        // ... rest of the login function remains the same ...
        // const response = await axios.post(`${API_BASE_URL}/users/login`, { username, password });
        // const { token, user } = response.data;
        // user.role = user.role || DEFAULT_ROLE; 
        
        localStorage.setItem('token', token);
        localStorage.setItem('user', JSON.stringify(user));
        setCurrentUser(user);
        setIsLoggedIn(true);
        axios.defaults.headers.common['Authorization'] = `Bearer ${token}`;
        return { success: true, message: `Login successful as ${role}!` };
    } catch (error) {
        const errorMessage = error.response?.data?.message || 'Login failed. Please check your credentials.';
        return { success: false, message: errorMessage };
    }
};
  // const login = async (username, password) => {
  //   console.log('AuthContext: Attempting login for user:', username);
  //   try {
  //     const response = await axios.post(`${API_BASE_URL}/users/login`, { username, password });
  //     const { token, user } = response.data;
  //     localStorage.setItem('token', token);
  //     localStorage.setItem('user', JSON.stringify(user));
  //     setCurrentUser(user);
  //     setIsLoggedIn(true); // Set state to true on successful login
  //     axios.defaults.headers.common['Authorization'] = `Bearer ${token}`;
  //     console.log('AuthContext: Login successful for user:', user.username);
  //     return { success: true, message: 'Login successful!' };
  //   } catch (error) {
  //     console.error('AuthContext: Login failed:', error);
  //     const errorMessage = error.response?.data?.message || 'Login failed. Please check your credentials.';
  //     return { success: false, message: errorMessage };
  //   }
  // };

  const register = async (username, password, role = 'general') => {
    console.log('AuthContext: Attempting registration for user:', username, 'with role:', role);
    try {
      const response = await axios.post(`${API_BASE_URL}/users/register`, { username, password, role });
      const { token, user } = response.data;
      localStorage.setItem('token', token);
      localStorage.setItem('user', JSON.stringify(user));
      setCurrentUser(user);
      setIsLoggedIn(true);
      axios.defaults.headers.common['Authorization'] = `Bearer ${token}`;
      console.log('AuthContext: Registration successful for user:', user.username);
      return { success: true, message: 'Registration successful!' };
    } catch (error) {
      console.error('AuthContext: Registration failed:', error);
      const errorMessage = error.response?.data?.message || 'Registration failed. Username might be taken or invalid.';
      return { success: false, message: errorMessage };
    }
  };

  const logout = () => {
    console.log('AuthContext: Logging out...');
    localStorage.removeItem('token');
    localStorage.removeItem('user');
    setIsLoggedIn(false);
    setCurrentUser(null);
    delete axios.defaults.headers.common['Authorization'];
    console.log('AuthContext: User logged out.');
  };

  return (
    <AuthContext.Provider value={{ isLoggedIn, currentUser, loading, login, register, logout }}>
      {children}
    </AuthContext.Provider>
  );
};

export const useAuth = () => useContext(AuthContext);
