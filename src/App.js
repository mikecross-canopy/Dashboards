import React, { useState, useEffect } from 'react';
import { Chart as ChartJS, ArcElement, Tooltip, Legend, CategoryScale, LinearScale, BarElement, Title } from 'chart.js';
import { Bar, Pie } from 'react-chartjs-2';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import './App.css';

// Register ChartJS components
ChartJS.register(ArcElement, Tooltip, Legend, CategoryScale, LinearScale, BarElement, Title);

// Initialize the Google Sheets document
const doc = new GoogleSpreadsheet('YOUR_GOOGLE_SHEET_ID');

function App() {
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [sheetData, setSheetData] = useState([]);

  useEffect(() => {
    const loadSheet = async () => {
      try {
        // You'll need to set up Google Sheets API credentials
        // For now, we'll use mock data
        // await doc.useServiceAccountAuth(require('./credentials.json'));
        // await doc.loadInfo();
        // const sheet = doc.sheetsByIndex[0];
        // const rows = await sheet.getRows();
        // setSheetData(rows);
        
        // Mock data for demonstration
        setTimeout(() => {
          setSheetData([
            { month: 'Jan', revenue: 12000, expenses: 8000 },
            { month: 'Feb', revenue: 15000, expenses: 9000 },
            { month: 'Mar', revenue: 18000, expenses: 10000 },
            { month: 'Apr', revenue: 16000, expenses: 8500 },
            { month: 'May', revenue: 21000, expenses: 11000 },
            { month: 'Jun', revenue: 23000, expenses: 12000 },
          ]);
          setLoading(false);
        }, 1000);
      } catch (err) {
        setError(err.message);
        setLoading(false);
      }
    };

    loadSheet();
  }, []);

  if (loading) return <div className="loading">Loading dashboard...</div>;
  if (error) return <div className="error">Error: {error}</div>;

  // Prepare data for charts
  const months = sheetData.map(item => item.month);
  const revenueData = sheetData.map(item => item.revenue);
  const expensesData = sheetData.map(item => item.expenses);

  const barData = {
    labels: months,
    datasets: [
      {
        label: 'Revenue',
        data: revenueData,
        backgroundColor: 'rgba(75, 192, 192, 0.6)',
        borderColor: 'rgba(75, 192, 192, 1)',
        borderWidth: 1,
      },
      {
        label: 'Expenses',
        data: expensesData,
        backgroundColor: 'rgba(255, 99, 132, 0.6)',
        borderColor: 'rgba(255, 99, 132, 1)',
        borderWidth: 1,
      },
    ],
  };

  const pieData = {
    labels: ['Revenue', 'Expenses'],
    datasets: [
      {
        data: [
          revenueData.reduce((a, b) => a + b, 0),
          expensesData.reduce((a, b) => a + b, 0),
        ],
        backgroundColor: [
          'rgba(75, 192, 192, 0.6)',
          'rgba(255, 99, 132, 0.6)',
        ],
        borderColor: [
          'rgba(75, 192, 192, 1)',
          'rgba(255, 99, 132, 1)',
        ],
        borderWidth: 1,
      },
    ],
  };

  const options = {
    responsive: true,
    plugins: {
      legend: {
        position: 'top',
      },
      title: {
        display: true,
        text: 'Monthly Financial Overview',
      },
    },
  };

  return (
    <div className="min-h-screen bg-gray-100 p-8">
      <header className="mb-8">
        <h1 className="text-3xl font-bold text-gray-800">Payment Analytics Dashboard</h1>
        <p className="text-gray-600">Executive Summary for Payments Account Management</p>
      </header>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        <div className="bg-white p-6 rounded-lg shadow-md">
          <h2 className="text-xl font-semibold mb-4">Monthly Revenue vs Expenses</h2>
          <div className="h-96">
            <Bar data={barData} options={options} />
          </div>
        </div>

        <div className="bg-white p-6 rounded-lg shadow-md">
          <h2 className="text-xl font-semibold mb-4">Revenue vs Expenses Distribution</h2>
          <div className="h-96">
            <Pie data={pieData} options={options} />
          </div>
        </div>

        <div className="lg:col-span-2 bg-white p-6 rounded-lg shadow-md">
          <h2 className="text-xl font-semibold mb-4">Transaction Data</h2>
          <div className="overflow-x-auto">
            <table className="min-w-full divide-y divide-gray-200">
              <thead className="bg-gray-50">
                <tr>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Month</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Revenue</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Expenses</th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Profit</th>
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-200">
                {sheetData.map((row, index) => (
                  <tr key={index} className={index % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                    <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">{row.month}</td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${row.revenue.toLocaleString()}</td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${row.expenses.toLocaleString()}</td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm font-medium">
                      <span className={row.revenue - row.expenses >= 0 ? 'text-green-600' : 'text-red-600'}>
                        ${(row.revenue - row.expenses).toLocaleString()}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      <footer className="mt-12 text-center text-gray-500 text-sm">
        <p>Last updated: {new Date().toLocaleString()}</p>
        <p className="mt-2">Data source: Google Sheets</p>
      </footer>
    </div>
  );
}

export default App;
