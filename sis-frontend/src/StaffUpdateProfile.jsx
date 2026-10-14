import React, { useState } from 'react';

const StaffUpdateProfile = ({ glassClass }) => {
    const [firstName, setFirstName] = useState('');
    const [lastName, setLastName] = useState('');
    const [email, setEmail] = useState('');
    const [message, setMessage] = useState('');

    const handleSubmit = (e) => {
        e.preventDefault();
        setMessage('This feature is under construction.');
        setFirstName('');
        setLastName('');
        setEmail('');
    };

    return (
        <div className={`p-4 rounded-xl ${glassClass}`}>
            <h2 className="text-2xl font-bold text-white mb-4">Update Profile</h2>
            {message && <p className="text-green-500">{message}</p>}
            <form onSubmit={handleSubmit} className="space-y-4">
                <div>
                    <label className="block text-sm font-medium text-gray-400">First Name:</label>
                    <input
                        type="text"
                        value={firstName}
                        onChange={(e) => setFirstName(e.target.value)}
                        className="w-full p-2 rounded-lg bg-gray-800 border border-gray-700 text-white"
                    />
                </div>
                <div>
                    <label className="block text-sm font-medium text-gray-400">Last Name:</label>
                    <input
                        type="text"
                        value={lastName}
                        onChange={(e) => setLastName(e.target.value)}
                        className="w-full p-2 rounded-lg bg-gray-800 border border-gray-700 text-white"
                    />
                </div>
                <div>
                    <label className="block text-sm font-medium text-gray-400">Email:</label>
                    <input
                        type="email"
                        value={email}
                        onChange={(e) => setEmail(e.target.value)}
                        className="w-full p-2 rounded-lg bg-gray-800 border border-gray-700 text-white"
                    />
                </div>
                <button type="submit" className="w-full bg-green-600 text-white font-bold py-2 rounded-lg transition duration-200">
                    Update Profile
                </button>
            </form>
        </div>
    );
};

export default StaffUpdateProfile;