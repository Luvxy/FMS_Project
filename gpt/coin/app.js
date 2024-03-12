const ctx = document.getElementById('gascoinChart').getContext('2d');
let chart;
const targetBuyPrice = 7000; // Example target buy price
const targetSellPrice = 10000; // Example target sell price

async function fetchDataAndUpdateChart() {
    const apiUrl = 'https://api.upbit.com/v1/candles/minutes/1?market=KRW-MASK&count=10'; // Adjusted for 1-minute candles

    try {
        const response = await fetch(apiUrl);
        const data = await response.json();
        // Ensure data is in chronological order
        data.reverse();

        // Filter data for the last 10 minutes (assuming data is already appropriate for this example)
        const prices = data.map(item => item.trade_price);
        const timestamps = data.map(item => new Date(item.timestamp).toLocaleTimeString());

        const formattedData = {
            labels: timestamps,
            datasets: [{
                label: 'MASK Price (KRW)',
                data: prices,
                borderColor: 'rgb(75, 192, 192)',
                tension: 0.1
            }, {
                label: 'Target Buy Price',
                data: Array(timestamps.length).fill(targetBuyPrice),
                borderColor: 'rgb(255, 99, 132)',
                borderWidth: 2,
                borderDash: [5, 5]
            }, {
                label: 'Target Sell Price',
                data: Array(timestamps.length).fill(targetSellPrice),
                borderColor: 'rgb(54, 162, 235)',
                borderWidth: 2,
                borderDash: [5, 5]
            }]
        };

        if (chart) {
            chart.data.labels = formattedData.labels;
            chart.data.datasets.forEach((dataset, i) => {
                dataset.data = formattedData.datasets[i].data;
            });
            chart.update();
        } else {
            chart = new Chart(ctx, {
                type: 'line',
                data: formattedData,
                options: {
                    scales: {
                        y: {
                            beginAtZero: false
                        }
                    }
                }
            });
        }
    } catch (error) {
        console.error('Failed to fetch data from Upbit API:', error);
    }
}

// Initial fetch and chart setup
fetchDataAndUpdateChart();

// Optionally, set an interval to update the chart in real-time (e.g., every minute)
setInterval(fetchDataAndUpdateChart, 600000);