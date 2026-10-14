import React, { useState, useEffect } from 'react';
import { AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, BarChart, Bar, Legend } from 'recharts';
import { DollarSign, Package, Users, TrendingUp, BarChart2, LayoutDashboard, ChevronsUp, ChevronsDown, Scale, HardHat, Banknote } from 'lucide-react';

// ----------------------------------------------------------------------
// --- 1. CORE CONSTANTS & HELPER COMPONENTS ---
// ----------------------------------------------------------------------

// Helper component to maintain consistent styling
const GlassCard = ({ title, value, icon: Icon, trend, children, className = '' }) => {
    const isPositive = trend >= 0;
    const TrendIcon = isPositive ? ChevronsUp : ChevronsDown;

    return (
        <div className={`p-4 rounded-xl bg-black/20 backdrop-blur-sm border border-white/10 flex flex-col justify-between h-full ${className}`}>
            <div className="flex items-start justify-between">
                <div className="text-white/80 font-semibold text-lg">{title}</div>
                <div className={`p-2 rounded-full ${isPositive ? 'bg-green-500/30 text-green-300' : 'bg-red-500/30 text-red-300'}`}>
                    <Icon className="w-5 h-5" />
                </div>
            </div>
            <div className="mt-4">
                {value && <div className="text-4xl font-extrabold text-white mb-2">
                    {value.toLocaleString('en-US', { minimumFractionDigits: 0 })}
                </div>}
                {trend !== undefined && (
                    <div className={`flex items-center text-sm ${isPositive ? 'text-green-400' : 'text-red-400'}`}>
                        <TrendIcon className="w-4 h-4 mr-1" />
                        {Math.abs(trend).toFixed(1)}% vs Last Period
                    </div>
                )}
                {children}
            </div>
        </div>
    );
};

// ----------------------------------------------------------------------
// --- 2. DATA TICKER HOOK ---
// ----------------------------------------------------------------------

const useDataTicker = (initialEmployeesLength = 35) => {
    const [tickingData, setTickingData] = useState(() => ({
        totalRevenue: 500000 + Math.floor(Math.random() * 100000),
        totalExpense: 200000 + Math.floor(Math.random() * 50000),
        totalAssets: 1500000 + Math.floor(Math.random() * 200000),
        activeEmployees: initialEmployeesLength,
        assetHistory: [1500000, 1510000, 1520000, 1530000, 1540000].map((v, i) => ({ month: `M${i+1}`, value: v })),
        revenueHistory: [400000, 450000, 500000, 550000, 600000].map((v, i) => ({ month: `M${i+1}`, value: v })),
        inventoryValue: 100000 + Math.floor(Math.random() * 20000),
    }));

    useEffect(() => {
        const interval = setInterval(() => {
            setTickingData(prev => {
                const volatilityFactor = 5;
                const newRev = Math.round(prev.totalRevenue + (Math.random() * 10000 * volatilityFactor - 5000 * volatilityFactor));
                const newExp = Math.round(prev.totalExpense + (Math.random() * 4000 * volatilityFactor - 2000 * volatilityFactor));
                const newAssets = Math.round(prev.totalAssets + (Math.random() * 20000 * volatilityFactor - 10000 * volatilityFactor));
                const newInv = Math.round(prev.inventoryValue + (Math.random() * 5000 * volatilityFactor - 2500 * volatilityFactor));

                // Update History (Shifting window)
                const newMonthIndex = parseInt(prev.assetHistory[prev.assetHistory.length - 1].month.substring(1)) + 1;
                const newAssetHistory = [...prev.assetHistory.slice(1), { month: `M${newMonthIndex}`, value: newAssets }];
                const newRevenueHistory = [...prev.revenueHistory.slice(1), { month: `M${newMonthIndex}`, value: newRev }];
                
                return {
                    totalRevenue: newRev,
                    totalExpense: newExp,
                    totalAssets: newAssets,
                    inventoryValue: newInv,
                    activeEmployees: prev.activeEmployees + (Math.random() > 0.9 ? (Math.random() > 0.5 ? 1 : -1) : 0),
                    assetHistory: newAssetHistory,
                    revenueHistory: newRevenueHistory,
                };
            });
        }, 3000); // 3 SECOND INTERVAL

        return () => clearInterval(interval);
    }, []);
    
    return tickingData;
};

// ----------------------------------------------------------------------
// --- 3. MOCK GRAPH DATA ---
// ----------------------------------------------------------------------

const MOCK_DEPT_PERFORMANCE = [
    { department: 'Finance', Revenue: 120000, Expense: 45000 },
    { department: 'HR', Revenue: 5000, Expense: 20000 },
    { department: 'Materials', Revenue: 15000, Expense: 55000 },
    { department: 'IT', Revenue: 1000, Expense: 25000 },
];
const MOCK_INVENTORY_STOCK = [
    { item: 'A4 Paper', stock: 1500 },
    { item: 'Laptop', stock: 25 },
    { item: 'Printer Toner', stock: 500 },
    { item: 'Cement', stock: 50 },
];

// ----------------------------------------------------------------------
// --- 4. MAIN DASHBOARD COMPONENT ---
// ----------------------------------------------------------------------

const IntelligentDashboard = ({ glassClass }) => {
    const data = useDataTicker(35);

    // Calculate profitability trend
    const currentProfit = data.totalRevenue - data.totalExpense;
    // Mock last period profit for trend calculation
    const lastPeriodProfit = (data.totalRevenue * 0.98) - (data.totalExpense * 1.05); 
    const profitTrend = ((currentProfit - lastPeriodProfit) / lastPeriodProfit) * 100;

    return (
        <div className="space-y-6">
            <h2 className="text-xl font-bold text-white/90 flex items-center mb-4">
                <LayoutDashboard className="w-5 h-5 mr-2 text-indigo-400" />
                Real-Time KPI Overview
            </h2>

            {/* Top Row: KPI Cards */}
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
                <GlassCard 
                    title="Total Revenue (ETB)" 
                    value={data.totalRevenue} 
                    icon={DollarSign} 
                    trend={Math.random() * 10 - 3} // Mock trend
                    className="col-span-1"
                />
                <GlassCard 
                    title="Total Assets (ETB)" 
                    value={data.totalAssets} 
                    icon={Scale} 
                    trend={Math.random() * 5 + 2} // Mock trend
                    className="col-span-1"
                />
                <GlassCard 
                    title="Active Employees" 
                    value={data.activeEmployees} 
                    icon={Users} 
                    trend={Math.random() * 1 - 0.5} // Mock trend
                    className="col-span-1"
                />
                 <GlassCard 
                    title="Profitability Trend" 
                    value={currentProfit} 
                    icon={Banknote} 
                    trend={profitTrend} 
                    className="col-span-1"
                />
            </div>
            
            {/* Second Row: Financial Charts */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                
                {/* 1. Revenue History Chart */}
                <div className={`lg:col-span-2 p-4 h-80 ${glassClass} flex flex-col`}>
                    <h3 className="font-semibold text-white/80 mb-4 flex items-center"><TrendingUp className="w-4 h-4 mr-2 text-green-400" /> Revenue & Expense Trend (Last 5 Periods)</h3>
                    <ResponsiveContainer width="100%" height="100%">
                        <AreaChart data={data.revenueHistory} margin={{ top: 10, right: 30, left: 0, bottom: 0 }}>
                            <CartesianGrid strokeDasharray="3 3" stroke="#ffffff20" />
                            <XAxis dataKey="month" stroke="#ffffff80" />
                            <YAxis stroke="#ffffff80" tickFormatter={(v) => `${(v / 1000).toFixed(0)}K`} />
                            <Tooltip contentStyle={{ backgroundColor: 'rgba(0,0,0,0.8)', border: 'none' }} formatter={(value) => [`${value.toLocaleString()} ETB`, 'Revenue']} />
                            <Area type="monotone" dataKey="value" stroke="#8884d8" fill="#8884d850" />
                            <Area type="monotone" dataKey={() => data.totalExpense} stroke="#f08080" fill="#f0808050" name="Expense" />
                        </AreaChart>
                    </ResponsiveContainer>
                </div>

                {/* 2. Inventory Value */}
                <div className={`p-4 h-80 ${glassClass} flex flex-col`}>
                    <h3 className="font-semibold text-white/80 mb-4 flex items-center"><Package className="w-4 h-4 mr-2 text-yellow-400" /> Inventory Valuation</h3>
                    <div className="flex-grow flex flex-col justify-center items-center">
                        <div className="text-5xl font-extrabold text-indigo-400">
                            {data.inventoryValue.toLocaleString('en-US', { minimumFractionDigits: 0 })}
                        </div>
                        <div className="text-white/60 mt-2">
                            Total Inventory Value (ETB)
                        </div>
                    </div>
                    <div className="text-xs text-white/50 border-t border-white/10 pt-2 mt-auto">
                        Value is automatically adjusted based on latest GRV/GIV records.
                    </div>
                </div>
            </div>

            {/* Third Row: Detailed Operations and Performance Charts */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                
                {/* 1. Departmental Performance */}
                <div className={`lg:col-span-2 p-4 h-80 ${glassClass} flex flex-col`}>
                    <h3 className="font-semibold text-white/80 mb-4 flex items-center"><BarChart2 className="w-4 h-4 mr-2 text-blue-400" /> Departmental Financial Performance</h3>
                    <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={MOCK_DEPT_PERFORMANCE} margin={{ top: 5, right: 30, left: 20, bottom: 5 }}>
                            <CartesianGrid strokeDasharray="3 3" stroke="#ffffff20" />
                            <XAxis dataKey="department" stroke="#ffffff80" />
                            <YAxis stroke="#ffffff80" tickFormatter={(v) => `${(v / 1000).toFixed(0)}K`} />
                            <Tooltip contentStyle={{ backgroundColor: 'rgba(0,0,0,0.8)', border: 'none' }} />
                            <Legend wrapperStyle={{ color: '#fff' }} />
                            <Bar dataKey="Revenue" fill="#4ade80" />
                            <Bar dataKey="Expense" fill="#fb7185" />
                        </BarChart>
                    </ResponsiveContainer>
                </div>
                
                {/* 2. Low Stock Alert */}
                <div className={`p-4 h-80 ${glassClass} flex flex-col`}>
                    <h3 className="font-semibold text-white/80 mb-4 flex items-center"><HardHat className="w-4 h-4 mr-2 text-red-400" /> Critical Stock Alert</h3>
                    <div className="space-y-3 overflow-y-auto flex-grow">
                        {MOCK_INVENTORY_STOCK.filter(i => i.stock < 1000).map((item, index) => (
                            <div key={index} className="flex justify-between items-center p-3 bg-red-900/40 rounded-lg border border-red-700/50">
                                <span className="font-medium text-red-300">{item.item}</span>
                                <span className="text-xl font-bold text-white">{item.stock}</span>
                            </div>
                        ))}
                    </div>
                    <div className="text-xs text-white/50 pt-2 mt-auto">
                        Items below re-order point (1000 units).
                    </div>
                </div>
            </div>
        </div>
    );
};

export default IntelligentDashboard;