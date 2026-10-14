import React from 'react';
import { useAuth } from '../AuthContext'; // Adjust path based on your structure
import { Link } from 'react-router-dom'; // Assuming react-router-dom is used for navigation

function Navbar() {
  const { currentUser, logout } = useAuth();

  return (
    <nav className="bg-gradient-to-r from-indigo-600 to-purple-700 p-4 shadow-lg">
      <div className="container mx-auto flex justify-between items-center">
        <Link to="/dashboard" className="text-white text-2xl font-bold rounded-md hover:bg-indigo-700 px-3 py-2 transition duration-300">
          SIS Dashboard
        </Link>
        <div className="flex items-center space-x-4">
          {currentUser && (
            <span className="text-white text-lg font-medium">
              Welcome, <span className="font-semibold">{currentUser.username}</span> ({currentUser.role})
            </span>
          )}
          <button
            onClick={logout}
            className="px-4 py-2 bg-white text-indigo-600 font-semibold rounded-md shadow-md hover:bg-gray-100 transition duration-300 ease-in-out"
          >
            Logout
          </button>
        </div>
      </div>
    </nav>
  );
}

export default Navbar;
