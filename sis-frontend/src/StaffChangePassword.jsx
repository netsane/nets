import React, { useState } from 'react';

const StaffChangePassword = ({ glassClass }) => {
    const [oldPassword, setOldPassword] = useState('');
    const [newPassword, setNewPassword] = useState('');
    const [confirmPassword, setConfirmPassword] = useState('');
    const [message, setMessage] = useState('');

    const handleSubmit = (e) => {
        e.preventDefault();
        setMessage('This feature is under construction.');
        setOldPassword('');
        setNewPassword('');
        setConfirmPassword('');
    };

    return (
        <div className={`p-4 rounded-xl ${glassClass}`}>
            <h2 className="text-2xl font-bold text-white mb-4">Change Password</h2>
            {message && <p className="text-green-500">{message}</p>}
            <form onSubmit={handleSubmit} className="space-y-4">
                <div>
                    <label className="block text-sm font-medium text-gray-400">Old Password:</label>
                    <input
                        type="password"
                        value={oldPassword}
                        onChange={(e) => setOldPassword(e.target.value)}
                        className="w-full p-2 rounded-lg bg-gray-800 border border-gray-700 text-white"
                    />
                </div>
                <div>
                    <label className="block text-sm font-medium text-gray-400">New Password:</label>
                    <input
                        type="password"
                        value={newPassword}
                        onChange={(e) => setNewPassword(e.target.value)}
                        className="w-full p-2 rounded-lg bg-gray-800 border border-gray-700 text-white"
                    />
                </div>
                <div>
                    <label className="block text-sm font-medium text-gray-400">Confirm New Password:</label>
                    <input
                        type="password"
                        value={confirmPassword}
                        onChange={(e) => setConfirmPassword(e.target.value)}
                        className="w-full p-2 rounded-lg bg-gray-800 border border-gray-700 text-white"
                    />
                </div>
                <button type="submit" className="w-full bg-green-600 text-white font-bold py-2 rounded-lg transition duration-200">
                    Change Password
                </button>
            </form>
        </div>
    );
};

export default StaffChangePassword;