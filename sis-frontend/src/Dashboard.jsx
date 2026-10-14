import React, { useState, useEffect, useRef } from "react";
import {
  Box, Typography, Divider, Button, Paper, AppBar, Toolbar, CircularProgress,
  List, ListItem, ListItemButton, ListItemText, ListItemIcon, IconButton,
  Accordion, AccordionSummary, AccordionDetails, Chip, Dialog, DialogTitle, DialogContent, DialogActions, TextField
} from "@mui/material";
import FavoriteIcon from '@mui/icons-material/Favorite';
import PeopleIcon from '@mui/icons-material/People';
import DeleteIcon from '@mui/icons-material/Delete';
import EditIcon from '@mui/icons-material/Edit';
import AddCircleOutlineIcon from "@mui/icons-material/AddCircleOutline";
import AccountCircleIcon from '@mui/icons-material/AccountCircle';
import ExpandMoreIcon from '@mui/icons-material/ExpandMore';
import ChevronLeftIcon from '@mui/icons-material/ChevronLeft';
import ChevronRightIcon from '@mui/icons-material/ChevronRight';
import VideoFileIcon from '@mui/icons-material/VideoFile';
import DescriptionIcon from '@mui/icons-material/Description';
import PhotoIcon from '@mui/icons-material/Photo';
import { useTheme } from '@mui/material/styles';
import axios from 'axios';
import { useAuth } from './AuthContext';
import { useNavigate } from 'react-router-dom';
import { initializeApp } from 'firebase/app';
import { getFirestore, doc, setDoc, deleteDoc, onSnapshot, collection } from 'firebase/firestore';
import { getAuth, signInWithCustomToken, onAuthStateChanged, signInAnonymously } from 'firebase/auth';

// ---- Logo Configuration ----
const UNIVERSITY_LOGO_PATH = '/images/logo.jpg';
const CAMPUS_PHOTO_PATH = '/images/campus_photo.jpg';

// --- API Base URL ---
const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || 'http://localhost:3001/api';

// --- WordPress API URL ---
const WORDPRESS_API_BASE_URL = '/api-proxy/graduation/wp-json/wp/v2';

// --- Helper function to format time ---
function formatTime(dt) {
  return dt.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
}

// --- Dedicated Image Slider Component with integrated history ---
function ImageSlider({ images, history }) {
  const [currentIndex, setCurrentIndex] = useState(0);

  // This useEffect hook resets the currentIndex to 0 whenever the images prop changes
  useEffect(() => {
    setCurrentIndex(0);
  }, [images]);

  const handleNext = () => {
    setCurrentIndex((prevIndex) => (prevIndex + 1) % images.length);
  };

  const handlePrev = () => {
    setCurrentIndex((prevIndex) => (prevIndex - 1 + images.length) % images.length);
  };

  if (!images || images.length === 0) {
    return (
      <Box sx={{
        position: 'relative',
        width: '100%',
        paddingBottom: '56.25%',
        borderRadius: '12px',
        overflow: 'hidden',
        boxShadow: '0 4px 12px rgba(0,0,0,0.3)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        bgcolor: 'rgba(255, 255, 255, 0.7)',
        color: '#1A2027',
      }}>
        <Typography variant="body1">No images available for this batch.</Typography>
      </Box>
    );
  }

  return (
    <Box sx={{
      position: 'relative',
      width: '100%',
      paddingBottom: '56.25%',
      borderRadius: '12px',
      overflow: 'hidden',
      boxShadow: '0 8px 16px rgba(0,0,0,0.4)',
      bgcolor: 'rgba(255, 255, 255, 0.7)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
    }}>
      <img
        src={images[currentIndex]}
        alt={`Slide ${currentIndex + 1}`}
        style={{
          position: 'absolute',
          top: 0,
          left: 0,
          width: '100%',
          height: '100%',
          objectFit: 'cover',
          transition: 'opacity 0.5s ease-in-out',
        }}
      />
      <IconButton
        onClick={handlePrev}
        sx={{
          position: 'absolute',
          left: 8,
          bgcolor: 'black',
          color: 'white',
          '&:hover': { bgcolor: 'rgba(0, 0, 0, 0.8)' },
          zIndex: 1,
        }}
      >
        <ChevronLeftIcon />
      </IconButton>
      <IconButton
        onClick={handleNext}
        sx={{
          position: 'absolute',
          right: 8,
          bgcolor: 'black',
          color: 'white',
          '&:hover': { bgcolor: 'rgba(0, 0, 0, 0.8)' },
          zIndex: 1,
        }}
      >
        <ChevronRightIcon />
      </IconButton>
      <Box sx={{
        position: 'absolute',
        bottom: 8,
        right: 8,
        bgcolor: 'rgba(0, 0, 0, 0.4)',
        color: 'white',
        px: 1,
        py: 0.5,
        borderRadius: '8px',
      }}>
        <Typography variant="body2">{`${currentIndex + 1} / ${images.length}`}</Typography>
      </Box>

      {/* Integrated Collapsible Student History */}
      {history && (
        <Box sx={{
          position: 'absolute',
          bottom: 0,
          left: 0,
          right: 0,
          zIndex: 2,
          bgcolor: 'rgba(255, 255, 255, 0.7)',
          borderRadius: '0 0 12px 12px',
          p: 2,
          backdropFilter: 'blur(5px)',
          transition: 'transform 0.3s ease-in-out',
        }}>
          <Accordion sx={{ bgcolor: 'transparent', boxShadow: 'none', '&:before': { display: 'none' } }}>
            <AccordionSummary
              expandIcon={<ExpandMoreIcon sx={{ color: '#1A2027' }} />}
              sx={{
                p: 0,
                '&.Mui-expanded': { minHeight: 48 },
                '& .MuiAccordionSummary-content.Mui-expanded': { mt: 1, mb: 1 }
              }}
            >
              <Typography variant="h5" sx={{ fontWeight: 'bold', color: '#1A2027' }}>
                Student History
              </Typography>
            </AccordionSummary>
            <AccordionDetails sx={{ bgcolor: 'transparent', p: 0, pt: 2 }}>
              <Typography variant="body1" sx={{ color: '#1A2027' }}>
                {history}
              </Typography>
            </AccordionDetails>
          </Accordion>
        </Box>
      )}
    </Box>
  );
}

function Dashboard() {
  // --- State Management ---
  const [batches, setBatches] = useState([]);
  const [resources, setResources] = useState({});
  const [selectedBatch, setSelectedBatch] = useState(null);
  const [newBatchName, setNewBatchName] = useState('');
  const [subscribed, setSubscribed] = useState({});
  const [likeState, setLikeState] = useState({});
  const [currentTime, setCurrentTime] = useState(new Date());
  const [wordpressPosts, setWordPressPosts] = useState([]);
  const [postUrl, setPostUrl] = useState(null);
  const [fetchError, setFetchError] = useState(null);

  const { currentUser, logout } = useAuth();
  const navigate = useNavigate();
  const theme = useTheme();

  // Document Management States
  const [openDocumentDialog, setOpenDocumentDialog] = useState(false);
  const [currentDocument, setCurrentDocument] = useState(null);
  const [documentTitle, setDocumentTitle] = useState('');
  const [documentUrl, setDocumentUrl] = useState('');
  const [documentError, setDocumentError] = useState('');

  // Confirmation Dialog States (for Delete Batch/Document)
  const [openConfirmDialog, setOpenConfirmDialog] = useState(false);
  const [confirmAction, setConfirmAction] = useState(null);
  const [confirmMessage, setConfirmMessage] = useState('');

  // Custom Message Dialog for user guidance
  const [openMessageDialog, setOpenMessageDialog] = useState(false);
  const [messageTitle, setMessageTitle] = useState('');
  const [messageContent, setMessageContent] = useState('');

  const selectedRes = resources[selectedBatch] || { id: null, images: [], docs: [], history: "No history available.", campusImage: '' };
  const selectedBatchObj = batches.find((b) => b.id === selectedBatch) || { id: null, name: "Select a Batch", likes: 0, subscriptions: 0 };


  // --- HANDLER FUNCTIONS ---
  const handleOpenWordPressPost = (url) => {
    if (url) {
      window.open(url, '_blank');
    }
  };

  const handleOpenMessageDialog = (title, content) => {
    setMessageTitle(title);
    setMessageContent(content);
    setOpenMessageDialog(true);
  };

  const handleCloseMessageDialog = () => {
    setOpenMessageDialog(false);
    setMessageTitle('');
    setMessageContent('');
  };

  const handleOpenAddDocumentDialog = () => {
    setCurrentDocument(null);
    setDocumentTitle('');
    setDocumentUrl('');
    setDocumentError('');
    setOpenDocumentDialog(true);
  };

  const handleOpenEditDocumentDialog = (doc) => {
    setCurrentDocument(doc);
    setDocumentTitle(doc.title);
    setDocumentUrl(doc.url);
    setDocumentError('');
    setOpenDocumentDialog(true);
  };

  const handleCloseDocumentDialog = () => {
    setOpenDocumentDialog(false);
    setCurrentDocument(null);
    setDocumentTitle('');
    setDocumentUrl('');
    setDocumentError('');
  };

  const handleSubmitDocument = async () => {
    if (!documentTitle || !documentUrl) {
      setDocumentError('Title and URL are required.');
      return;
    }

    setDocumentError('');

    try {
      const token = localStorage.getItem('token');
      const config = { headers: { Authorization: `Bearer ${token}` } };

      if (currentDocument) {
        // Mock update
        const updatedDoc = { ...currentDocument, title: documentTitle, url: documentUrl };
        setResources(prev => ({
          ...prev,
          [selectedBatch]: {
            ...prev[selectedBatch],
            docs: prev[selectedBatch].docs.map(doc =>
              doc.id === updatedDoc.id ? updatedDoc : doc
            )
          }
        }));
      } else {
        const newDoc = {
          id: Date.now().toString(),
          resource_id: selectedBatch,
          title: documentTitle,
          url: documentUrl,
          type: documentUrl.match(/\.(jpeg|jpg|png|gif|webp|svg)$/i) ? 'photo' : (documentUrl.match(/(youtube\.com|youtu\.be)/i) ? 'video' : 'document'),
        };
        setResources(prev => ({
          ...prev,
          [selectedBatch]: {
            ...prev[selectedBatch],
            docs: [...(prev[selectedBatch]?.docs || []), newDoc]
          }
        }));
      }
      handleCloseDocumentDialog();
    } catch (error) {
      console.error('Error saving document:', error.response?.data || error.message);
      setDocumentError(error.response?.data?.message || 'Failed to save document.');
      if (error.response && (error.response.status === 401 || error.response.status === 403)) {
        logout();
      }
    }
  };

  const confirmDeleteBatch = () => {
    if (selectedBatch === null || batches.length === 0) return;
    const batchToDelete = batches.find(b => b.id === selectedBatch);
    if (!batchToDelete) return;

    setConfirmMessage(`Are you sure you want to delete batch "${batchToDelete.name}"? This action cannot be undone.`);
    setConfirmAction(() => async () => {
      try {
        const updatedBatches = batches.filter(b => b.id !== selectedBatch);
        setBatches(updatedBatches);
        setResources(prevResources => {
          const newResources = { ...prevResources };
          delete newResources[selectedBatch];
          return newResources;
        });

        setSelectedBatch(updatedBatches.length > 0 ? updatedBatches[0].id : null);
        handleCloseConfirmDialog();
      } catch (e) {
        console.error("Error deleting batch:", e);
        setConfirmMessage(`Failed to delete batch: ${e.response?.data?.message || e.message}`);
        if (e.response && (e.response.status === 401 || e.response.status === 403)) {
          logout();
        }
      }
    });
    setOpenConfirmDialog(true);
  };

  const confirmDeleteDocument = (documentId) => {
    setConfirmMessage('Are you sure you want to delete this document? This action cannot be undone.');
    setConfirmAction(() => async () => {
      try {
        setResources(prev => ({
          ...prev,
          [selectedBatch]: {
            ...prev[selectedBatch],
            docs: prev[selectedBatch].docs.filter(doc => doc.id !== documentId)
          }
        }));
        handleCloseConfirmDialog();
      } catch (error) {
        console.error('Error deleting document:', error.response?.data || error.message);
        setConfirmMessage(`Failed to delete document: ${error.response?.data?.message || error.message}`);
        if (error.response && (error.response.status === 401 || error.response.status === 403)) {
          logout();
        }
      }
    });
    setOpenConfirmDialog(true);
  };

  const handleCloseConfirmDialog = () => {
    setOpenConfirmDialog(false);
    setConfirmAction(null);
    setConfirmMessage('');
  };

  const handleAddBatch = () => {
    const name = newBatchName.trim();
    if (!name) return;

    const newBatchId = (batches.length + 1).toString();
    const newBatchData = {
      id: newBatchId,
      name,
      likes: 0,
      subscriptions: 0
    };

    setBatches(prevBatches => [...prevBatches, newBatchData].sort((a, b) => parseInt(a.id) - parseInt(b.id)));
    setResources(prevResources => ({
      ...prevResources,
      [newBatchId]: getResourcesForBatch(newBatchId)
    }));
    setSelectedBatch(newBatchId);
    setNewBatchName("");
  };

  const handleLike = (batchId) => {
    if (likeState[batchId]) return;

    setBatches(prevBatches =>
      prevBatches.map(b =>
        b.id === batchId ? { ...b, likes: b.likes + 1 } : b
      )
    );
    setLikeState(prev => ({ ...prev, [batchId]: true }));
  };

  const handleSubscribe = (batchId) => {
    setSubscribed((prev) => ({
      ...prev,
      [batchId]: !prev[batchId]
    }));
    setBatches(prevBatches =>
      prevBatches.map(b =>
        b.id === batchId
          ? { ...b, subscriptions: subscribed[batchId] ? b.subscriptions - 1 : b.subscriptions + 1 }
          : b
      )
    );
  };

  // --- MOCK DATA FUNCTIONS ---
  // This function is now responsible for providing the mock data
  const getResourcesForBatch = (batchId) => {
    if (String(batchId) === '43') {
      return {
        id: '43',
        history: "Photos and videos from the 43rd batch commencement ceremony on August 02, 2025.",
        images: [
          'uploaded:20250802_085131.jpg-0a81eec4-40df-439c-b720-fc5a3320dd98',
          'uploaded:20250802_091409.jpg-0df410f1-3007-42a5-afa4-0906778b1cf6',
          'uploaded:20250802_093924.jpg-58056055-9bd4-41cb-906c-45a888db71b6',
          'uploaded:20250802_095046.jpg-ab9fd93d-f10a-4606-9193-37356c999055',
          'uploaded:20250802_095114.jpg-674c4819-0945-4476-a6c2-af7b2f29785e',
          'uploaded:20250802_095118.jpg-40e0a64c-7f40-4cbc-8a7a-ebeb3b6dcc96',
          'uploaded:20250802_101514.jpg-08be4ddf-2bdb-4431-b5a9-f7e74b5f94d1',
          'uploaded:20250802_101517(0).jpg-5aba4478-70bf-4949-a254-761f972b7e31',
          'uploaded:20250802_101517.jpg-19a67152-72f0-44fd-896f-3273922ea502',
          'uploaded:20250802_101533.jpg-ec7812b5-59ce-4f3b-a377-88a164cb02f9',
        ],
        docs: [
          { id: '1', title: "Graduation Video", url: "https://www.youtube.com/watch?v=dQw4w9WgXcQ", type: "video" },
          { id: '2', title: "Batch Report 2023", url: "https://example.com/documents/report_2023.pdf", type: "document" }
        ],
      };
    }
    return {
      id: batchId,
      history: `No detailed history available for this batch.`,
      images: [],
      docs: [],
    };
  };

  // --- EFFECTS ---
  useEffect(() => {
    const interval = setInterval(() => setCurrentTime(new Date()), 1000);
    return () => clearInterval(interval);
  }, []);

  useEffect(() => {
    const fetchAllBatches = () => {
      const fetchedBatches = [
        { id: '43', name: 'Batch 43 Commencement', likes: 25, subscriptions: 10 },
        { id: '2', name: 'Another Batch', likes: 10, subscriptions: 5 },
        { id: '3', name: 'Third Batch', likes: 15, subscriptions: 8 },
      ];
      setBatches(fetchedBatches);
      if (fetchedBatches.length > 0) {
        setSelectedBatch(String(fetchedBatches[0].id));
      }
    };
    fetchAllBatches();
  }, []);

  useEffect(() => {
    if (selectedBatch) {
      const fetchedResources = getResourcesForBatch(selectedBatch);
      setResources(prev => ({
        ...prev,
        [selectedBatch]: fetchedResources
      }));
    }
  }, [selectedBatch]);

  useEffect(() => {
    const fetchWordPressPosts = async () => {
      try {
        const response = await axios.get(`${WORDPRESS_API_BASE_URL}/posts?_embed`);
        if (response.data && response.data.length > 0) {
          setWordPressPosts(response.data);
          setPostUrl(response.data[0].link);
        } else {
          setWordPressPosts([]);
          setPostUrl(null);
        }
      } catch (error) {
        console.error("Error fetching WordPress posts:", error);
        setWordPressPosts([]);
        setPostUrl(null);
      }
    };
    fetchWordPressPosts();
  }, []);

  // Reusable 3D button style
  const threeDButtonStyle = {
    position: 'relative',
    overflow: 'hidden',
    fontWeight: 'bold',
    borderRadius: '12px',
    padding: '10px 20px',
    transition: 'transform 0.1s, box-shadow 0.1s',
    '&:hover': {
      transform: 'translateY(-2px)',
      boxShadow: '0 4px 6px rgba(0, 0, 0, 0.4), 0 1px 1px rgba(0, 0, 0, 0.15)',
    },
    '&:active': {
      transform: 'translateY(1px)',
      boxShadow: '0 1px 2px rgba(0, 0, 0, 0.2)',
    }
  };

  return (
    <Box sx={{
      height: "100vh",
      width: "100%",
      display: "flex",
      flexDirection: "column",
      overflow: "hidden",
      backgroundImage: `url(${CAMPUS_PHOTO_PATH})`,
      backgroundSize: 'cover',
      backgroundPosition: 'center',
      backgroundAttachment: 'fixed',
      position: 'relative',
      '@keyframes blinking': {
        '0%': { boxShadow: '0 0 8px 2px rgba(255, 255, 255, 0.7)' },
        '50%': { boxShadow: '0 0 12px 4px rgba(255, 255, 255, 0.2)' },
        '100%': { boxShadow: '0 0 8px 2px rgba(255, 255, 255, 0.7)' },
      },
    }}>
      {/* Main Content Container */}
      <Box sx={{
        position: 'relative',
        zIndex: 1,
        display: "flex",
        flexDirection: "column",
        flexGrow: 1,
        maxHeight: '100vh',
        bgcolor: 'rgba(0, 0, 0, 0.05)',
        borderRadius: '12px',
        m: 2,
        boxShadow: '0 0 15px rgba(0,0,0,0.2)'
      }}>
        {/* AppBar/Header */}
        <AppBar position="static" sx={{
          bgcolor: 'transparent',
          boxShadow: 'none',
          color: '#ffffff',
        }}>
          <Toolbar sx={{ justifyContent: 'space-between' }}>
            <Box sx={{ display: 'flex', alignItems: 'center' }}>
              <Box sx={{ height: 40, width: 40, marginRight: 2, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                <img src={UNIVERSITY_LOGO_PATH} alt="University Logo" style={{ height: '100%', width: '100%' }} />
              </Box>
              <Typography variant="h5" sx={{ fontWeight: 'bold', textShadow: '2px 2px 4px rgba(0,0,0,0.5)' }}>
                SIS Dashboard
              </Typography>
            </Box>
            <Box sx={{ display: 'flex', alignItems: 'center' }}>
              <Button
                onClick={() => handleOpenWordPressPost(postUrl)}
                disabled={!postUrl}
                sx={{
                  ...threeDButtonStyle,
                  bgcolor: 'transparent',
                  border: '1px solid',
                  borderColor: 'rgba(255, 255, 255, 0.5)',
                  color: 'white',
                  mr: 2,
                }}
              >
                Go to WordPress Post
              </Button>
              <Typography variant="body1" sx={{ mr: 2, textShadow: '1px 1px 2px rgba(0,0,0,0.5)' }}>
                {currentTime.toLocaleDateString()}
              </Typography>
              <Typography variant="body1" sx={{ mr: 2, textShadow: '1px 1px 2px rgba(0,0,0,0.5)' }}>
                {formatTime(currentTime)}
              </Typography>
              <Button onClick={logout} endIcon={<PeopleIcon />} sx={{
                ...threeDButtonStyle,
                color: 'white',
                bgcolor: 'transparent',
                border: '1px solid',
                borderColor: 'rgba(255, 255, 255, 0.5)',
                animation: 'blinking 1.5s infinite ease-in-out',
              }}
              >
                Logout
              </Button>
            </Box>
          </Toolbar>
        </AppBar>
        {/* Main Section */}
        <Box sx={{ display: 'flex', flexGrow: 1, overflow: 'auto' }}>
          {/* Left Sidebar */}
          <Paper elevation={0} sx={{
            width: 280,
            p: 2,
            overflow: 'auto',
            flexShrink: 0,
            bgcolor: 'rgba(255, 255, 255, 0.05)',
            color: '#ffffff',
            boxShadow: 'none'
          }}>
            <Typography variant="h6" gutterBottom sx={{ fontWeight: 'bold', textShadow: '1px 1px 2px rgba(0,0,0,0.5)' }}>
              Batches
            </Typography>
            <Divider sx={{ mb: 2, bgcolor: 'rgba(255, 255, 255, 0.1)' }} />
            <Box sx={{ mb: 2, display: 'flex' }}>
              <TextField
                size="small"
                label="New Batch"
                variant="outlined"
                value={newBatchName}
                onChange={(e) => setNewBatchName(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && handleAddBatch()}
                sx={{
                  flexGrow: 1,
                  mr: 1,
                  bgcolor: 'transparent',
                  input: { color: 'white' },
                  '& label': { color: 'rgba(255, 255, 255, 0.7)' },
                  '& fieldset': { borderColor: 'rgba(255, 255, 255, 0.3)' }
                }}
              />
              <Button
                variant="contained"
                onClick={handleAddBatch}
                sx={{
                  ...threeDButtonStyle,
                  bgcolor: 'transparent',
                  border: '1px solid',
                  borderColor: 'primary.main',
                  '&:hover': {
                    bgcolor: 'rgba(25, 118, 210, 0.2)',
                  },
                }}
              >
                Add
              </Button>
            </Box>
            <List dense>
              {Array.isArray(batches) && batches.length > 0 ? (
                batches.map((batch) => (
                  <ListItem key={batch.id} disablePadding>
                    <ListItemButton
                      selected={selectedBatch === batch.id}
                      onClick={() => setSelectedBatch(batch.id)}
                      sx={{
                        borderRadius: '8px',
                        '&.Mui-selected': {
                          bgcolor: 'rgba(255, 255, 255, 0.1)',
                          '&:hover': {
                            bgcolor: 'rgba(255, 255, 255, 0.15)',
                          },
                        },
                      }}
                    >
                      <ListItemText
                        primary={batch.name}
                        primaryTypographyProps={{
                          fontWeight: 'medium',
                          sx: { textShadow: '1px 1px 2px rgba(0,0,0,0.5)' }
                        }}
                      />
                    </ListItemButton>
                  </ListItem>
                ))
              ) : (
                <Typography variant="body2" sx={{ color: 'rgba(255, 255, 255, 0.7)', mt: 2 }}>
                  No batches available.
                </Typography>
              )}
            </List>
            <Divider sx={{ my: 2, bgcolor: 'rgba(255, 255, 255, 0.1)' }} />
            <Box>
              <Typography variant="h6" gutterBottom sx={{ fontWeight: 'bold', textShadow: '1px 1px 2px rgba(0,0,0,0.5)' }}>
                WordPress Posts
              </Typography>
              {wordpressPosts.length > 0 ? (
                <List dense>
                  {wordpressPosts.map((post) => (
                    <ListItem key={post.id} disablePadding>
                      <ListItemButton
                        onClick={() => handleOpenWordPressPost(post.link)}
                        sx={{
                          borderRadius: '8px',
                          '&:hover': {
                            bgcolor: 'rgba(255, 255, 255, 0.15)',
                          },
                        }}
                      >
                        <ListItemText
                          primary={post.title.rendered}
                          primaryTypographyProps={{
                            fontWeight: 'medium',
                            sx: { textShadow: '1px 1px 2px rgba(0,0,0,0.5)' }
                          }}
                        />
                      </ListItemButton>
                    </ListItem>
                  ))}
                </List>
              ) : (
                <Typography variant="body2" sx={{ color: 'rgba(255, 255, 255, 0.7)' }}>
                  No WordPress posts available.
                </Typography>
              )}
            </Box>
          </Paper>

          {/* Right Main Content Area */}
          <Box sx={{ flexGrow: 1, p: 3, overflow: 'auto' }}>
            {selectedBatch ? (
              <Box>
                <Paper elevation={3} sx={{
                  p: 3,
                  mb: 3,
                  bgcolor: 'rgba(255, 255, 255, 0.05)',
                  boxShadow: '0 4px 8px rgba(0,0,0,0.2)',
                  color: 'white'
                }}>
                  <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', mb: 2 }}>
                    <Typography variant="h4" sx={{ fontWeight: 'bold', textShadow: '2px 2px 4px rgba(0,0,0,0.5)' }}>
                      {selectedBatchObj.name}
                    </Typography>
                    <Box>
                      <Button
                        variant="contained"
                        onClick={() => handleLike(selectedBatch)}
                        startIcon={<FavoriteIcon />}
                        disabled={likeState[selectedBatch]}
                        sx={{
                          ...threeDButtonStyle,
                          bgcolor: 'transparent',
                          border: '1px solid',
                          borderColor: theme.palette.primary.main,
                          color: 'white',
                          mr: 1,
                        }}
                      >
                        {selectedBatchObj.likes}
                      </Button>
                      <Button
                        variant="contained"
                        onClick={() => handleSubscribe(selectedBatch)}
                        startIcon={<PeopleIcon />}
                        sx={{
                          ...threeDButtonStyle,
                          bgcolor: 'transparent',
                          border: '1px solid',
                          borderColor: theme.palette.primary.main,
                          color: 'white',
                        }}
                      >
                        {selectedBatchObj.subscriptions}
                      </Button>
                    </Box>
                  </Box>
                  <Box sx={{ display: 'flex', gap: 2, mb: 2 }}>
                    <Button
                      variant="contained"
                      onClick={handleOpenAddDocumentDialog}
                      startIcon={<AddCircleOutlineIcon />}
                      sx={{
                        ...threeDButtonStyle,
                        bgcolor: 'transparent',
                        border: '1px solid',
                        borderColor: 'secondary.main',
                        color: 'white',
                      }}
                    >
                      Add Document
                    </Button>
                    <Button
                      variant="contained"
                      onClick={confirmDeleteBatch}
                      startIcon={<DeleteIcon />}
                      sx={{
                        ...threeDButtonStyle,
                        bgcolor: 'transparent',
                        border: '1px solid',
                        borderColor: theme.palette.error.main,
                        color: 'white',
                      }}
                    >
                      Delete Batch
                    </Button>
                  </Box>

                  <Typography variant="body1" sx={{ mt: 2, mb: 2, p: 2, borderRadius: '8px', bgcolor: 'rgba(255, 255, 255, 0.1)' }}>
                    {selectedRes.history}
                  </Typography>

                  <Box sx={{ mt: 3 }}>
                    <ImageSlider images={selectedRes.images} history={selectedRes.history} />
                  </Box>

                  {/* Documents Section */}
                  <Box sx={{ mt: 4 }}>
                    <Typography variant="h5" sx={{ fontWeight: 'bold', mb: 2, textShadow: '1px 1px 2px rgba(0,0,0,0.5)' }}>
                      Documents
                    </Typography>
                    {selectedRes.docs && selectedRes.docs.length > 0 ? (
                      <List dense>
                        {selectedRes.docs.map(doc => (
                          <ListItem
                            key={doc.id}
                            secondaryAction={
                              <Box>
                                <IconButton edge="end" aria-label="edit" onClick={() => handleOpenEditDocumentDialog(doc)}>
                                  <EditIcon sx={{ color: 'white' }} />
                                </IconButton>
                                <IconButton edge="end" aria-label="delete" onClick={() => confirmDeleteDocument(doc.id)}>
                                  <DeleteIcon sx={{ color: 'white' }} />
                                </IconButton>
                              </Box>
                            }
                          >
                            <ListItemButton
                              component="a"
                              href={doc.url}
                              target="_blank"
                              sx={{
                                color: 'white',
                                '&:hover': {
                                  bgcolor: 'rgba(255, 255, 255, 0.1)',
                                },
                              }}
                            >
                              <ListItemIcon>
                                {doc.type === "photo" ? (
                                  <PhotoIcon sx={{ color: '#E3A832' }} />
                                ) : doc.type === "video" ? (
                                  <VideoFileIcon sx={{ color: '#E3A832' }} />
                                ) : (
                                  <DescriptionIcon sx={{ color: '#54724B' }} />
                                )}
                              </ListItemIcon>
                              <ListItemText primary={doc.title} sx={{ color: 'white' }} />
                            </ListItemButton>
                          </ListItem>
                        ))}
                      </List>
                    ) : (
                      <Typography variant="body2" sx={{ color: 'rgba(255, 255, 255, 0.7)' }}>
                        No documents available for this batch.
                      </Typography>
                    )}
                  </Box>

                </Paper>
              </Box>
            ) : (
              <Box sx={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '100%' }}>
                <Typography variant="h6" sx={{ color: 'rgba(255, 255, 255, 0.7)' }}>
                  Please select or add a batch to view its resources.
                </Typography>
              </Box>
            )}
          </Box>
        </Box>
      </Box>

      {/* Confirmation Dialog */}
      <Dialog open={openConfirmDialog} onClose={handleCloseConfirmDialog} PaperProps={{ sx: {
        bgcolor: 'rgba(255, 255, 255, 0.05)',
        boxShadow: '0 4px 8px rgba(0,0,0,0.2)',
        color: 'white'
      }}}>
        <DialogTitle>
          <Typography sx={{ color: 'white' }}>Confirm Action</Typography>
        </DialogTitle>
        <DialogContent>
          <Typography sx={{ color: 'white' }}>{confirmMessage}</Typography>
        </DialogContent>
        <DialogActions>
          <Button
            onClick={handleCloseConfirmDialog}
            sx={{
              ...threeDButtonStyle,
              bgcolor: 'transparent',
              border: '1px solid',
              borderColor: theme.palette.primary.main,
              color: 'white',
            }}
          >
            Cancel
          </Button>
          <Button
            onClick={confirmAction}
            autoFocus
            sx={{
              ...threeDButtonStyle,
              bgcolor: 'transparent',
              border: '1px solid',
              borderColor: theme.palette.error.main,
              color: 'white',
            }}
          >
            Confirm
          </Button>
        </DialogActions>
      </Dialog>

      {/* Custom Message Dialog for user feedback */}
      <Dialog open={openMessageDialog} onClose={handleCloseMessageDialog} PaperProps={{ sx: {
        bgcolor: 'rgba(255, 255, 255, 0.05)',
        boxShadow: '0 4px 8px rgba(0,0,0,0.2)',
        color: 'white'
      }}}>
          <DialogTitle>
            <Typography sx={{ color: 'white' }}>{messageTitle}</Typography>
          </DialogTitle>
          <DialogContent>
              <Typography sx={{ color: 'white' }}>{messageContent}</Typography>
          </DialogContent>
          <DialogActions>
              <Button
                onClick={handleCloseMessageDialog}
                variant="contained"
                sx={{
                  ...threeDButtonStyle,
                  bgcolor: 'transparent',
                  border: '1px solid',
                  borderColor: theme.palette.primary.main,
                  color: 'white',
                }}
              >
                OK
              </Button>
          </DialogActions>
      </Dialog>

      {/* Document Management Dialog */}
      <Dialog open={openDocumentDialog} onClose={handleCloseDocumentDialog} PaperProps={{ sx: {
        bgcolor: 'rgba(255, 255, 255, 0.05)',
        boxShadow: '0 4px 8px rgba(0,0,0,0.2)',
        color: 'white',
      }}}>
        <DialogTitle sx={{ color: 'white' }}>{currentDocument ? 'Edit Document' : 'Add New Document'}</DialogTitle>
        <DialogContent>
          <Typography color="error" variant="body2" sx={{ mb: 2 }}>{documentError}</Typography>
          <TextField
            autoFocus
            margin="dense"
            label="Document Title"
            type="text"
            fullWidth
            variant="outlined"
            value={documentTitle}
            onChange={(e) => setDocumentTitle(e.target.value)}
            sx={{
              mb: 2,
              input: { color: 'white' },
              '& label': { color: 'rgba(255, 255, 255, 0.7)' },
              '& fieldset': { borderColor: 'rgba(255, 255, 255, 0.3)' },
            }}
          />
          <TextField
            margin="dense"
            label="Document URL"
            type="url"
            fullWidth
            variant="outlined"
            value={documentUrl}
            onChange={(e) => setDocumentUrl(e.target.value)}
            sx={{
              input: { color: 'white' },
              '& label': { color: 'rgba(255, 255, 255, 0.7)' },
              '& fieldset': { borderColor: 'rgba(255, 255, 255, 0.3)' },
            }}
          />
        </DialogContent>
        <DialogActions>
          <Button
            onClick={handleCloseDocumentDialog}
            sx={{
              ...threeDButtonStyle,
              bgcolor: 'transparent',
              border: '1px solid',
              borderColor: 'theme.palette.primary.main',
              color: 'white',
            }}
          >
            Cancel
          </Button>
          <Button
            onClick={handleSubmitDocument}
            sx={{
              ...threeDButtonStyle,
              bgcolor: 'transparent',
              border: '1px solid',
              borderColor: 'theme.palette.secondary.main',
              color: 'white',
            }}
          >
            {currentDocument ? 'Save Changes' : 'Add Document'}
          </Button>
        </DialogActions>
      </Dialog>
    </Box>
  );
}

export default Dashboard;
