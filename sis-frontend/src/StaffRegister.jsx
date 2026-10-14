import React, { useState } from 'react';
import { useAuth } from './AuthContext';

const StaffRegister = ({ glassClass }) => {
    const { register } = useAuth();
    const [username, setUsername] = useState('');
    const [password, setPassword] = useState('');
    const [role, setRole] = useState('employee');
    const [message, setMessage] = useState('');

    const handleSubmit = async (e) => {
        e.preventDefault();
        setMessage('');

        const result = await register(username, password, role);
        if (result.success) {
            setMessage(result.message);
        } else {
            setMessage(result.message);
        }
    };

    return (
        <div className={`p-4 rounded-xl ${glassClass}`}>
            <h2 className="text-2xl font-bold text-white mb-4">Staff Registration</h2>
            {message && <p className="text-green-500">{message}</p>}
            <form onSubmit={handleSubmit} className="space-y-4">
                <div>
                    <label className="block text-sm font-medium text-gray-400">Username:</label>
                    <input
                        type="text"
                        value={username}
                        onChange={(e) => setUsername(e.target.value)}
                        className="w-full p-2 rounded-lg bg-gray-800 border border-gray-700 text-white"
                    />
                </div>
                <div>
                    <label className="block text-sm font-medium text-gray-400">Password:</label>
                    <input
                        type="password"
                        value={password}
                        onChange={(e) => setPassword(e.target.value)}
                        className="w-full p-2 rounded-lg bg-gray-800 border border-gray-700 text-white"
                    />
                </div>
                <div>
                    <label className="block text-sm font-medium text-gray-400">Role:</label>
                    <select
                        value={role}
                        onChange={(e) => setRole(e.target.value)}
                        className="w-full p-2 rounded-lg bg-gray-800 border border-gray-700 text-white"
                    >
                        <option value="employee">Employee</option>
                        <option value="hr_staff">HR Staff</option>
                        <option value="hr_manager">HR Manager</option>
                        <option value="admin">Admin</option>
                    </select>
                </div>
                <button type="submit" className="w-full bg-green-600 text-white font-bold py-2 rounded-lg transition duration-200">
                    Register
                </button>
            </form>
        </div>
    );
};

export default StaffRegister;