<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Payments Analytics Dashboard</title>
    <link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>📊</text></svg>">
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-zoom@2.0.1"></script>
    <script src="https://accounts.google.com/gsi/client" async defer></script>
    <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">
    <link rel="stylesheet" href="styles.css">
</head>
<body class="bg-gray-100">
    <div class="container mx-auto px-4 py-8">
        <header class="mb-8">
            <h1 class="text-4xl font-bold mb-3" style="background: linear-gradient(135deg, #2B4BFF 0%, #0BE6C7 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text;">Payments Analytics Dashboard</h1>
            <p class="text-lg font-semibold" style="color: #000000;"></p>
        </header>

        <!-- Date Filter Controls -->
        <div class="bg-white p-6 rounded-lg shadow mb-8">
            <div class="flex flex-col lg:flex-row lg:items-start lg:justify-between gap-6 mb-6">
                <div>
                    <h3 class="text-lg font-semibold text-gray-800 mb-2">Global View Filter</h3>
                    <div class="flex items-center space-x-2">
                        <label class="text-sm font-medium text-gray-700">Show:</label>
                        <div class="flex bg-gray-100 rounded-lg p-1">
                            <button id="globalClosedWonToggle" class="px-4 py-2 text-sm rounded-md bg-blue-500 text-white transition-colors">Closed Won Only</button>
                            <button id="globalTotalOppsToggle" class="px-4 py-2 text-sm rounded-md text-gray-700 hover:bg-gray-200 transition-colors">All Opportunities</button>
                            <button id="globalPipelineToggle" class="px-4 py-2 text-sm rounded-md text-gray-700 hover:bg-gray-200 transition-colors">Pipeline</button>
                        </div>
                    </div>
                </div>
                <div class="flex-1 max-w-md">
                    <h3 class="text-lg font-semibold text-gray-800 mb-4">User Filter</h3>
                    <div class="mb-4">
                        <label class="block text-sm font-medium text-gray-700 mb-2">Select User to Display</label>
                        <select id="userFilterSelect" multiple size="6" class="w-full border border-gray-300 rounded-md px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500">
                            <!-- Options will be populated by JavaScript -->
                        </select>
                        <p class="text-xs text-gray-500 mt-1">Select a specific user or "All Users" to view everyone</p>
                    </div>
                </div>
                <div class="flex-1 max-w-md">
                    <h3 class="text-lg font-semibold text-gray-800 mb-4">Date Range Filter</h3>
                    <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                        <div>
                            <label class="block text-sm font-medium text-gray-700 mb-2">Quick Filters</label>
                            <select id="dateFilterSelect" class="border border-gray-300 rounded-md px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500" style="min-width: 160px;">
                                <option value="all-time" selected>All Time</option>
                                <option value="from-2025-04-01">From 4/1/2025</option>
                                <option value="this-month">This Month</option>
                                <option value="last-month">Last Month</option>
                                <option value="this-quarter">This Quarter</option>
                                <option value="last-quarter">Last Quarter</option>
                                <option value="this-year">This Year</option>
                                <option value="custom">Custom Range</option>
                            </select>
                        </div>
                        <div id="customDateRange" class="hidden">
                            <label class="block text-sm font-medium text-gray-700 mb-2">Start Date</label>
                            <input type="date" id="startDateInput" class="w-full border border-gray-300 rounded-md px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500">
                        </div>
                        <div id="customDateRangeEnd" class="hidden">
                            <label class="block text-sm font-medium text-gray-700 mb-2">End Date</label>
                            <input type="date" id="endDateInput" class="w-full border border-gray-300 rounded-md px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500">
                        </div>
                        
                    </div>
                </div>
                <div class="flex items-end lg:items-center">
                    <button id="toggleActivitiesBtn" class="bg-gray-800 hover:bg-gray-900 text-white font-medium py-2 px-4 rounded-md text-sm transition duration-200">Show Activities</button>
                </div>
            </div>
        </div>

        <div class="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-8">
            <div class="bg-white p-6 rounded-lg shadow">
                <h2 class="text-xl font-semibold mb-4">Monthly GPV Trend</h2>
                <div class="chart-container">
                    <canvas id="volumeChart"></canvas>
                </div>
            </div>
            <div class="bg-white p-6 rounded-lg shadow">
                <h2 class="text-xl font-semibold mb-4">Deal Size Distribution</h2>
                <div class="chart-container">
                    <canvas id="paymentMethodsChart"></canvas>
                </div>
            </div>
            <div class="bg-white p-6 rounded-lg shadow">
                <h2 class="text-xl font-semibold mb-4">Monthly Deal Count</h2>
                <div class="chart-container">
                    <canvas id="successRateChart"></canvas>
                </div>
            </div>
            <div class="bg-white p-6 rounded-lg shadow">
                <h2 class="text-xl font-semibold mb-4">Average Deal Size by Month</h2>
                <div class="chart-container">
                    <canvas id="revenueByRegionChart"></canvas>
                </div>
            </div>
        </div>

        <!-- Additional Analytics Charts -->
        <div class="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-8">
            <div class="bg-white p-6 rounded-lg shadow">
                <h2 class="text-xl font-semibold mb-4">Conversion Funnel</h2>
                <div class="chart-container">
                    <canvas id="funnelChart"></canvas>
                </div>
            </div>
            <div class="bg-white p-6 rounded-lg shadow overflow-hidden">
                <h2 class="text-xl font-semibold mb-4">Top Performers Leaderboard</h2>
                <div class="overflow-x-auto" style="max-height: 320px;">
                    <table class="min-w-full divide-y divide-gray-200" id="leaderboardTable">
                        <thead class="bg-gray-50">
                            <tr>
                                <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Owner</th>
                                <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">GPV</th>
                                <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Deals</th>
                                <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Win Rate</th>
                            </tr>
                        </thead>
                        <tbody class="bg-white divide-y divide-gray-200" id="leaderboardBody">
                            <!-- Populated by JS -->
                        </tbody>
                    </table>
                </div>
            </div>
            <div class="bg-white p-6 rounded-lg shadow">
                <h2 class="text-xl font-semibold mb-4">Performance by Opportunity Owner</h2>
                <div class="chart-container">
                    <canvas id="ownerPerformanceChart"></canvas>
                </div>
            </div>
            <div class="bg-white p-6 rounded-lg shadow">
                <h2 class="text-xl font-semibold mb-4">Win Rate by Opportunity Owner (Rolling 6 Months)</h2>
                <div class="chart-container">
                    <canvas id="winRateChart"></canvas>
                </div>
            </div>
            <div class="bg-white p-6 rounded-lg shadow">
                <h2 class="text-xl font-semibold mb-4">Opportunity Source Analysis</h2>
                <div class="chart-container">
                    <canvas id="sourceChart"></canvas>
                </div>
            </div>
            <div class="bg-white p-6 rounded-lg shadow">
                <h2 class="text-xl font-semibold mb-4">Deal Age Distribution</h2>
                <div class="chart-container">
                    <canvas id="ageChart"></canvas>
                </div>
            </div>
            <div class="bg-white p-6 rounded-lg shadow">
                <h2 class="text-xl font-semibold mb-4">Current Payment Processors</h2>
                <div class="chart-container">
                    <canvas id="processorChart"></canvas>
                </div>
            </div>
            <div class="bg-white p-6 rounded-lg shadow">
                <h2 class="text-xl font-semibold mb-4">GPV Segment Distribution</h2>
                <div class="chart-container">
                    <canvas id="segmentChart"></canvas>
                </div>
            </div>
        </div>

        <!-- Owner Source Analysis Section -->
        <div class="bg-white p-6 rounded-lg shadow mb-8">
            <div class="flex justify-between items-center mb-4">
                <div>
                    <h2 class="text-xl font-semibold">Opportunity Owner Source Analysis</h2>
                    <p class="text-sm text-gray-600">Shows each owner's opportunities by source (GPV and quantity)</p>
                </div>
                <div class="flex items-center space-x-2">
                    <label class="text-sm font-medium text-gray-700">View:</label>
                    <div class="flex bg-gray-100 rounded-lg p-1">
                        <button id="closedWonToggle" class="px-3 py-1 text-sm rounded-md bg-blue-500 text-white transition-colors">Closed Won</button>
                        <button id="totalOppsToggle" class="px-3 py-1 text-sm rounded-md text-gray-700 hover:bg-gray-200 transition-colors">Total Opportunities</button>
                    </div>
                </div>
            </div>
            <p id="ownerSourceTop3" class="text-sm text-gray-600 mb-3"></p>
            <div class="chart-container">
                <canvas id="ownerSourceChart" style="height: 600px;"></canvas>
            </div>
        </div>

        <!-- Team Activities - Matt, Carson, Jacob, Jaxon & Nicholas -->
        <div id="activitiesSection" class="bg-white p-6 rounded-lg shadow mb-8 hidden">
            <div class="flex justify-between items-center mb-4">
                <div>
                    <h2 class="text-xl font-semibold">Payments AE Activity</h2>
                    <p class="text-sm text-gray-600">Activity metrics and demo performance for the core team members</p>
                </div>
                <div class="flex-1 max-w-md ml-4">
                    <h3 class="text-lg font-semibold text-gray-800 mb-2">Activities Date Range Filter</h3>
                    <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-2">
                        <div>
                            <label class="block text-xs font-medium text-gray-700 mb-1">Quick Filters</label>
                            <select id="activitiesDateFilterSelect" class="w-full border border-gray-300 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                                <option value="all-time">All Time</option>
                                <option value="from-2025-01-01">From 1/1/2025</option>
                                <option value="this-year">This Year</option>
                                <option value="this-quarter">This Quarter</option>
                                <option value="last-quarter">Last Quarter</option>
                                <option value="this-month">This Month</option>
                                <option value="last-month">Last Month</option>
                                <option value="custom">Custom Range</option>
                            </select>
                        </div>
                        <div id="activitiesCustomDateRange" class="hidden">
                            <label class="block text-xs font-medium text-gray-700 mb-1">Start Date</label>
                            <input type="date" id="activitiesStartDateInput" class="w-full border border-gray-300 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                        </div>
                        <div id="activitiesCustomDateRangeEnd" class="hidden">
                            <label class="block text-xs font-medium text-gray-700 mb-1">End Date</label>
                            <input type="date" id="activitiesEndDateInput" class="w-full border border-gray-300 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                        </div>
                        <div class="flex items-end">
                            <button id="activitiesApplyFilterBtn" class="bg-blue-500 hover:bg-blue-600 text-white font-medium py-1 px-3 rounded-md text-sm transition duration-200">
                                Apply
                            </button>
                        </div>
                    </div>
                </div>
            </div>
            <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-6" id="activities-metrics">
                <!-- Metrics will be populated by JavaScript -->
            </div>
            <div class="chart-container">
                <canvas id="activitiesChart"></canvas>
            </div>
            <div class="mt-6">
                <h3 class="text-lg font-semibold mb-2">AE Breakdown</h3>
                <div class="overflow-x-auto" style="max-height: 360px; overflow-y: auto;">
                    <table id="activitiesBreakdownTable" class="min-w-full divide-y divide-gray-200">
                        <thead>
                            <tr>
                                <th class="px-6 py-3 bg-gray-50 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">AE</th>
                                <th class="px-6 py-3 bg-gray-50 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Calls</th>
                                <th class="px-6 py-3 bg-gray-50 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Emails</th>
                                <th class="px-6 py-3 bg-gray-50 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Demo Sets</th>
                                <th class="px-6 py-3 bg-gray-50 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Answered %</th>
                                <th class="px-6 py-3 bg-gray-50 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Calls/Demo</th>
                                <th class="px-6 py-3 bg-gray-50 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Emails/Demo</th>
                                <th class="px-6 py-3 bg-gray-50 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Avg Demo Call</th>
                            </tr>
                        </thead>
                        <tbody id="activitiesBreakdownTableBody" class="bg-white divide-y divide-gray-200"></tbody>
                    </table>
                </div>
            </div>
        </div>

        

    <script src="config.public.js"></script>
    <script src="app.js"></script>
</body>
</html>
