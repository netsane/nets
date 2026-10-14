import React, { useState, useEffect, useCallback } from 'react';
import { RefreshCw, Globe, Server, AlertTriangle } from 'lucide-react';

// The main application component
const App = () => {
    // State to hold the fetched data
    const [data, setData] = useState(null);
    // State to track the loading status
    const [isLoading, setIsLoading] = useState(false);
    // State for any error messages
    const [error, setError] = useState(null);
    // State for the currently selected API target key ('test' or 'myuniversity')
    const [target, setTarget] = useState('myuniversity'); 

    // Base URL for the proxy server you just confirmed is working
    // NOTE: This must match the port your Node server is running on (3001)
    const PROXY_BASE_URL = 'http://localhost:3001/api/external/';

    // The function to fetch data from the proxy
    const fetchData = useCallback(async (currentTarget) => {
        setIsLoading(true);
        setError(null);
        setData(null);

        const url = `${PROXY_BASE_URL}${currentTarget}`;

        try {
            const response = await fetch(url);

            if (!response.ok) {
                // Try to read JSON error message from the server
                const errorBody = await response.json().catch(() => ({ message: 'Unknown server error' }));
                throw new Error(`HTTP Error ${response.status}: ${errorBody.error || errorBody.message}`);
            }

            const result = await response.json();
            setData(result);

        } catch (err) {
            console.error('Fetch error:', err);
            setError(`Could not fetch data from Proxy. Make sure your Node server is running on port 3001. Details: ${err.message}`);
        } finally {
            setIsLoading(false);
        }
    }, []);

    // Effect to run the fetch every time the 'target' state changes
    useEffect(() => {
        // Automatically fetch data for the initial target
        fetchData(target);
    }, [target, fetchData]); 

    // Helper function to change the target and trigger a re-fetch
    const handleTargetChange = (newTarget) => {
        if (newTarget !== target) {
            setTarget(newTarget);
        }
    };

    // Determine target details for display
    const targetInfo = {
        'test': { name: 'Example.com', url: 'http://example.com' },
        'myuniversity': { name: 'MIT.edu (Your Live Target)', url: 'https://www.mit.edu/' },
    }[target];

    return (
        <div className="min-h-screen bg-gray-50 p-4 sm:p-8 font-sans">
            <script src="https://cdn.tailwindcss.com"></script>
            <style>{`
                @import url('https://fonts.googleapis.com/css2?family=Inter:wght@100..900&display=swap');
                body { font-family: 'Inter', sans-serif; }
            `}</style>
            
            <div className="max-w-4xl mx-auto">
                <header className="text-center mb-10 p-6 bg-white shadow-xl rounded-xl">
                    <h1 className="text-4xl font-extrabold text-blue-800 mb-2">CORS Proxy Frontend</h1>
                    <p className="text-gray-600">
                        Consuming data from your backend proxy running on <code className="bg-gray-100 p-1 rounded font-mono text-sm">http://localhost:3001</code>
                    </p>
                </header>

                {/* Target Selector */}
                <div className="flex justify-center space-x-4 mb-8">
                    <button
                        onClick={() => handleTargetChange('myuniversity')}
                        disabled={isLoading}
                        className={`px-6 py-3 rounded-xl font-semibold transition duration-300 shadow-md flex items-center ${
                            target === 'myuniversity' 
                                ? 'bg-indigo-600 text-white hover:bg-indigo-700' 
                                : 'bg-white text-indigo-600 border border-indigo-200 hover:bg-indigo-50'
                        } disabled:opacity-50`}
                    >
                        <Globe className="w-5 h-5 mr-2" />
                        Live Target (MIT.edu)
                    </button>
                    <button
                        onClick={() => handleTargetChange('test')}
                        disabled={isLoading}
                        className={`px-6 py-3 rounded-xl font-semibold transition duration-300 shadow-md flex items-center ${
                            target === 'test' 
                                ? 'bg-indigo-600 text-white hover:bg-indigo-700' 
                                : 'bg-white text-indigo-600 border border-indigo-200 hover:bg-indigo-50'
                        } disabled:opacity-50`}
                    >
                        <Server className="w-5 h-5 mr-2" />
                        Test Target (Example.com)
                    </button>
                </div>

                {/* Status Bar */}
                <div className="bg-white p-4 mb-6 rounded-xl shadow-lg border-l-4 border-indigo-500">
                    <p className="font-medium text-gray-700">
                        Current Target: <span className="font-mono text-indigo-600">{targetInfo.name}</span>
                        <br/>
                        Requesting URL: <span className="text-sm font-mono text-gray-500">{PROXY_BASE_URL}{target}</span>
                    </p>
                </div>

                {/* Content Area: Loading, Error, or Data */}
                <div className="min-h-96 bg-white p-6 rounded-xl shadow-2xl">
                    
                    {/* 1. Loading State */}
                    {isLoading && (
                        <div className="flex flex-col items-center justify-center h-full text-indigo-600 p-10">
                            <RefreshCw className="w-8 h-8 animate-spin mb-3" />
                            <p className="text-lg font-medium">Fetching and Scraping...</p>
                            <p className="text-sm text-gray-500 mt-1">Calling your Node.js proxy now.</p>
                        </div>
                    )}

                    {/* 2. Error State */}
                    {error && (
                        <div className="bg-red-50 border-l-4 border-red-500 text-red-700 p-4 rounded-lg flex items-start" role="alert">
                            <AlertTriangle className="w-5 h-5 mr-3 mt-1 flex-shrink-0" />
                            <div>
                                <p className="font-bold">API Error</p>
                                <p className="text-sm font-mono whitespace-pre-wrap">{error}</p>
                            </div>
                        </div>
                    )}

                    {/* 3. Data Display */}
                    {data && data.length > 0 && !isLoading && (
                        <>
                            <h2 className="text-2xl font-bold mb-4 text-gray-800 border-b pb-2">Scraped Content ({data.length} Items)</h2>
                            {data.map((item, index) => (
                                <div key={index} className="mb-6 p-4 border border-gray-100 rounded-lg bg-gray-50">
                                    <p className="text-sm text-gray-500 mb-1">
                                        Source: <span className="font-semibold">{item.source}</span>
                                    </p>
                                    <h3 className="text-xl font-bold text-blue-700 mb-2">{item.title}</h3>
                                    <p className="text-gray-700 leading-relaxed italic">{item.content}</p>
                                    <a 
                                        href={item.url} 
                                        target="_blank" 
                                        rel="noopener noreferrer" 
                                        className="text-sm text-indigo-500 hover:text-indigo-700 mt-2 inline-block"
                                    >
                                        View Original Page
                                    </a>
                                </div>
                            ))}
                        </>
                    )}

                    {/* 4. No Data/Empty State */}
                    {data && data.length === 0 && !isLoading && (
                        <div className="text-center p-10 text-gray-500">
                            <p className="text-lg font-medium">No content was successfully scraped.</p>
                            <p className="text-sm mt-1">Check your proxy server logs and Cheerio selectors for the `{target}` target.</p>
                        </div>
                    )}
                </div>
            </div>
        </div>
    );
};

export default App;
