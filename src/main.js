
import './style.css'
import Chart from 'chart.js/auto'
import * as XLSX from 'xlsx'

const workbookUrl = '/data/Superstore.xlsx'
const sheetName = 'Sample - Superstore'

document.querySelector('#app').innerHTML = `
  <div class="dashboard">
    <aside class="sidebar">
      <div class="brand">
        <div class="brand-icon">R</div>
        <div>
          <h2>RetailBI</h2>
          <span>Business Intelligence</span>
        </div>
      </div>

      <p class="menu-label">MENU UTAMA</p>
      <nav>
        <a class="nav-item active" href="#">▦ Dashboard Overview</a>
        <a class="nav-item" href="#sales">▤ Sales Analytics</a>
        <a class="nav-item" href="#products">▣ Product Performance</a>
      </nav>

      <div class="sidebar-bottom">
        <span>RETAIL SALES INTELLIGENCE</span>
        <p>Business Analytics Platform</p>
      </div>
    </aside>

    <main class="main-content">
      <header class="topbar">
        <div>
          <p class="breadcrumb">Pages / Dashboard</p>
          <h1>Dashboard Overview</h1>
        </div>
        <div class="profile">
          <span class="status-dot"></span>
          <span>BI Dashboard</span>
        </div>
      </header>

      <section class="welcome">
        <div>
          <p class="eyebrow">BUSINESS PERFORMANCE</p>
          <h2>Retail Sales Intelligence</h2>
          <p>Monitor your sales, understand performance, and discover business insights.</p>
        </div>
        <div class="welcome-icon">📊</div>
      </section>

      <section class="section-heading">
        <div>
          <h2>Business Overview</h2>
          <p>Key performance indicators</p>
        </div>
        <span class="demo-label">LOADING DATA...</span>
      </section>

      <section class="filter-panel">
        <div class="filter-title">
          <h3>Filter Data</h3>
          <p>Customize your business analysis</p>
        </div>

        <div class="filter-controls">
          <div class="filter-group">
            <label for="yearFilter">Year</label>
            <select id="yearFilter">
              <option value="all">All Years</option>
            </select>
          </div>

          <div class="filter-group">
            <label for="categoryFilter">Category</label>
            <select id="categoryFilter">
              <option value="all">All Categories</option>
            </select>
          </div>
        </div>
      </section>

      <section class="kpi-grid">
        <article class="kpi-card">
          <div class="kpi-top">
            <span>Total Sales</span>
            <span class="kpi-icon">💰</span>
          </div>
          <h3 id="totalSales">Loading...</h3>
          <p>Actual dataset value</p>
        </article>

        <article class="kpi-card">
          <div class="kpi-top">
            <span>Total Profit</span>
            <span class="kpi-icon">📈</span>
          </div>
          <h3 id="totalProfit">Loading...</h3>
          <p>Actual dataset value</p>
        </article>

        <article class="kpi-card">
          <div class="kpi-top">
            <span>Total Orders</span>
            <span class="kpi-icon">🛒</span>
          </div>
          <h3 id="totalOrders">Loading...</h3>
          <p>Actual dataset value</p>
        </article>

        <article class="kpi-card">
          <div class="kpi-top">
            <span>Products</span>
            <span class="kpi-icon">📦</span>
          </div>
          <h3 id="totalProducts">Loading...</h3>
          <p>Actual dataset value</p>
        </article>
      </section>

      <section class="charts-grid" id="sales">
        <article class="panel sales-panel">
          <div class="panel-heading">
            <div>
              <h3>Sales Performance</h3>
              <p>Monthly sales overview</p>
            </div>
          </div>
          <div class="chart-container">
            <canvas id="salesChart"></canvas>
          </div>
        </article>

        <article class="panel">
          <div class="panel-heading">
            <div>
              <h3>Sales by Category</h3>
              <p>Category distribution</p>
            </div>
          </div>
          <div class="chart-container">
            <canvas id="categoryChart"></canvas>
          </div>
        </article>
      </section>

      <section class="panel product-panel" id="products">
        <div class="panel-heading">
          <div>
            <h3>Business Insights</h3>
            <p>Analytical observations from the dataset</p>
          </div>
        </div>

        <div id="insights" class="insight-placeholder">
          <span>💡</span>
          <div>
            <strong>Analyzing business data...</strong>
            <p>Insights will appear after data processing.</p>
          </div>
        </div>
      </section>

      <section class="panel product-panel" id="product-performance">
        <div class="panel-heading">
          <div>
            <h3>Top 10 Products Performance</h3>
            <p>Products ranked by total sales</p>
          </div>
        </div>

        <div class="table-wrapper">
          <table class="product-table">
            <thead>
              <tr>
                <th>Product Name</th>
                <th>Category</th>
                <th>Total Sales</th>
                <th>Total Profit</th>
              </tr>
            </thead>
            <tbody id="productTableBody">
              <tr>
                <td colspan="4">Loading product data...</td>
              </tr>
            </tbody>
          </table>
        </div>
      </section>

      <footer>RetailBI · Retail Sales Intelligence Dashboard</footer>
    </main>
  </div>
`

let allRows = []
let salesChart = null
let categoryChart = null

const money = value =>
  '$' + Number(value || 0).toLocaleString('en-US', {
    maximumFractionDigits: 0
  })

const number = value =>
  Number(value || 0).toLocaleString('en-US')

function parseOrderDate(value) {
  if (!value) return null

  if (value instanceof Date && !isNaN(value.getTime())) {
    return value
  }

  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value)

    if (!parsed) return null

    return new Date(parsed.y, parsed.m - 1, parsed.d)
  }

  const text = String(value).trim()
  const usDate = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/)

  if (usDate) {
    return new Date(
      Number(usDate[3]),
      Number(usDate[1]) - 1,
      Number(usDate[2])
    )
  }

  const parsedDate = new Date(text)

  return isNaN(parsedDate.getTime()) ? null : parsedDate
}

function populateFilters(rows) {
  const yearFilter = document.getElementById('yearFilter')
  const categoryFilter = document.getElementById('categoryFilter')

  const years = [...new Set(
    rows.map(row => row.orderDate.getFullYear())
  )].sort((a, b) => a - b)

  const categories = [...new Set(
    rows.map(row => row.Category).filter(Boolean)
  )].sort()

  years.forEach(year => {
    const option = document.createElement('option')
    option.value = String(year)
    option.textContent = year
    yearFilter.appendChild(option)
  })

  categories.forEach(category => {
    const option = document.createElement('option')
    option.value = category
    option.textContent = category
    categoryFilter.appendChild(option)
  })

  yearFilter.addEventListener('change', updateDashboard)
  categoryFilter.addEventListener('change', updateDashboard)
}

function updateDashboard() {
  const selectedYear = document.getElementById('yearFilter').value
  const selectedCategory = document.getElementById('categoryFilter').value

  const filteredRows = allRows.filter(row => {
    const yearMatch =
      selectedYear === 'all' ||
      String(row.orderDate.getFullYear()) === selectedYear

    const categoryMatch =
      selectedCategory === 'all' ||
      row.Category === selectedCategory

    return yearMatch && categoryMatch
  })

  updateKPIs(filteredRows)
  updateCharts(filteredRows)
  updateInsights(filteredRows)
  updateProductTable(filteredRows)

  document.querySelector('.demo-label').textContent =
    `${filteredRows.length.toLocaleString('en-US')} RECORDS`
}

function updateKPIs(rows) {
  const totalSales = rows.reduce((sum, row) => sum + row.Sales, 0)
  const totalProfit = rows.reduce((sum, row) => sum + row.Profit, 0)

  const totalOrders = new Set(
    rows.map(row => row['Order ID']).filter(Boolean)
  ).size

  const totalProducts = new Set(
    rows.map(row => row['Product Name']).filter(Boolean)
  ).size

  document.getElementById('totalSales').textContent = money(totalSales)
  document.getElementById('totalProfit').textContent = money(totalProfit)
  document.getElementById('totalOrders').textContent = number(totalOrders)
  document.getElementById('totalProducts').textContent = number(totalProducts)
}

function updateCharts(rows) {
  const monthlySales = {}
  const categorySales = {}

  rows.forEach(row => {
    const date = row.orderDate
    const monthKey =
      `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`

    monthlySales[monthKey] =
      (monthlySales[monthKey] || 0) + row.Sales

    const category = row.Category || 'Other'

    categorySales[category] =
      (categorySales[category] || 0) + row.Sales
  })

  const months = Object.keys(monthlySales).sort()
  const categories = Object.keys(categorySales)

  if (salesChart) salesChart.destroy()
  if (categoryChart) categoryChart.destroy()

  salesChart = new Chart(document.getElementById('salesChart'), {
    type: 'line',
    data: {
      labels: months,
      datasets: [{
        label: 'Sales',
        data: months.map(month => monthlySales[month]),
        borderColor: '#5367e8',
        backgroundColor: 'rgba(83, 103, 232, 0.12)',
        fill: true,
        tension: 0.4,
        borderWidth: 3
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'top' }
      },
      scales: {
        y: {
          ticks: {
            callback: value => '$' + Number(value).toLocaleString('en-US')
          }
        }
      }
    }
  })

  categoryChart = new Chart(document.getElementById('categoryChart'), {
    type: 'doughnut',
    data: {
      labels: categories,
      datasets: [{
        data: categories.map(category => categorySales[category]),
        backgroundColor: ['#5367e8', '#20b99a', '#f2b84b', '#ed718a'],
        borderWidth: 0
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom' }
      }
    }
  })
}

function updateInsights(rows) {
  const container = document.getElementById('insights')

  if (!rows.length) {
    container.innerHTML = `
      <span>⚠️</span>
      <div>
        <strong>No data available</strong>
        <p>There are no records matching the selected filters.</p>
      </div>
    `
    return
  }

  const totalSales = rows.reduce((sum, row) => sum + row.Sales, 0)
  const totalProfit = rows.reduce((sum, row) => sum + row.Profit, 0)

  const categorySales = {}
  const productSales = {}

  rows.forEach(row => {
    const category = row.Category || 'Other'
    const product = row['Product Name'] || 'Unknown'

    categorySales[category] =
      (categorySales[category] || 0) + row.Sales

    productSales[product] =
      (productSales[product] || 0) + row.Sales
  })

  const bestCategory = Object.entries(categorySales)
    .sort((a, b) => b[1] - a[1])[0]

  const bestProduct = Object.entries(productSales)
    .sort((a, b) => b[1] - a[1])[0]

  const profitMargin = totalSales
    ? (totalProfit / totalSales) * 100
    : 0

  container.innerHTML = `
    <span>💡</span>
    <div>
      <strong>Business Performance Analysis</strong>
      <p>Total sales mencapai ${money(totalSales)}.</p>
      <p>Total profit yang diperoleh adalah ${money(totalProfit)}.</p>
      <p>Kategori dengan penjualan tertinggi adalah
        <strong>${bestCategory[0]}</strong>
        sebesar ${money(bestCategory[1])}.
      </p>
      <p>Produk dengan penjualan tertinggi adalah
        <strong>${bestProduct[0]}</strong>
        sebesar ${money(bestProduct[1])}.
      </p>
      <p>Profit margin keseluruhan adalah ${profitMargin.toFixed(2)}%.</p>
    </div>
  `
}

function updateProductTable(rows) {
  const tableBody = document.getElementById('productTableBody')

  if (!rows.length) {
    tableBody.innerHTML = `
      <tr>
        <td colspan="4">No products found for the selected filters.</td>
      </tr>
    `
    return
  }

  const products = {}

  rows.forEach(row => {
    const productName = row['Product Name'] || 'Unknown Product'
    const category = row.Category || 'Other'

    if (!products[productName]) {
      products[productName] = {
        name: productName,
        category,
        sales: 0,
        profit: 0
      }
    }

    products[productName].sales += row.Sales
    products[productName].profit += row.Profit
  })

  const topProducts = Object.values(products)
    .sort((a, b) => b.sales - a.sales)
    .slice(0, 10)

  tableBody.innerHTML = topProducts.map((product, index) => `
    <tr>
      <td>
        <span class="rank-number">${index + 1}</span>
        ${product.name}
      </td>
      <td>${product.category}</td>
      <td>${money(product.sales)}</td>
      <td>${money(product.profit)}</td>
    </tr>
  `).join('')
}

async function loadExcelData() {
  try {
    const response = await fetch(workbookUrl)

    if (!response.ok) {
      throw new Error(
        'File Excel tidak ditemukan. Periksa folder public/data.'
      )
    }

    const buffer = await response.arrayBuffer()

    const workbook = XLSX.read(buffer, {
      type: 'array',
      cellDates: true
    })

    const sheet = workbook.Sheets[sheetName]

    if (!sheet) {
      throw new Error(`Sheet ${sheetName} tidak ditemukan.`)
    }

    const rows = XLSX.utils.sheet_to_json(sheet, {
      defval: null
    })

    if (!rows.length) {
      throw new Error('Data Excel kosong.')
    }

    allRows = rows.map(row => ({
      ...row,
      Sales: Number(row.Sales) || 0,
      Profit: Number(row.Profit) || 0,
      orderDate: parseOrderDate(row['Order Date'])
    })).filter(row => row.orderDate !== null)

    if (!allRows.length) {
      throw new Error('Data tanggal Order Date tidak dapat dibaca.')
    }

    populateFilters(allRows)
    updateDashboard()

    console.log('Dashboard berhasil menggunakan data Excel!')
    console.log('Jumlah data:', allRows.length)
    console.log('Nama sheet:', sheetName)

  } catch (error) {
    console.error('Gagal membaca Excel:', error)

    document.querySelector('.demo-label').textContent = 'DATA ERROR'

    document.getElementById('insights').innerHTML = `
      <span>⚠️</span>
      <div>
        <strong>Data gagal dimuat</strong>
        <p>${error.message}</p>
      </div>
    `
  }
}

loadExcelData()
