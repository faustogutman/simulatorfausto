<h1>🏡 Mortgage and Property Comparison Calculator</h1>

<p>
This Python application provides a robust and user-friendly tool for comparing different properties and their associated mortgage scenarios. Built with <strong>Tkinter</strong> for the graphical user interface and <strong>pandas</strong> for data handling and calculations, it allows users to input property details, explore various loan scenarios, and export detailed reports to Excel.
</p>

<h2>✨ Features</h2>

<ul>
  <li><strong>Multi-Property Comparison:</strong> Manage and analyze up to three different properties simultaneously using an intuitive tabbed interface.</li>
  <li><strong>Detailed Property Inputs:</strong> Input essential data for each property, including:
    <ul>
      <li>Alias & Link</li>
      <li>Property Price</li>
      <li>Area (Square Meters)</li>
      <li>Loan-to-Value (LTV) Percentage</li>
      <li>Expected Monthly Rent</li>
      <li>Options to skip Purchase Tax and Broker Fees</li>
    </ul>
  </li>
  <li><strong>Multiple Loan Scenarios:</strong> For each property, define up to three distinct loan scenarios by specifying:
    <ul>
      <li>Annual Interest Rate</li>
      <li>Loan Term (in years)</li>
    </ul>
  </li>
  <li><strong>Comprehensive Calculations:</strong> The application automatically calculates:
    <ul>
      <li>Estimated Purchase Tax</li>
      <li>Required Down Payment</li>
      <li>Estimated Lawyer Fees</li>
      <li>Estimated Broker Fees</li>
      <li>Total Capital Required</li>
      <li>Price per Square Meter</li>
      <li>Loan Amount</li>
      <li>Monthly Mortgage Payments for each scenario</li>
      <li>Total Interest and Total Payment over the loan term for each scenario</li>
    </ul>
  </li>
  <li><strong>Amortization Tables & Graphs:</strong> Visualize the breakdown of principal and interest payments over time for each loan scenario with detailed amortization tables and interactive plots.</li>
  <li><strong>Rent-to-Mortgage Comparison:</strong> See a clear ratio of expected monthly rent to the calculated monthly mortgage payment for each scenario.</li>
  <li><strong>Data Persistence (CSV):</strong>
    <ul>
      <li><strong>Save Inputs:</strong> Export all property input data to a CSV file for easy saving and sharing.</li>
      <li><strong>Load Inputs:</strong> Import previously saved CSV data to quickly populate your property tabs.</li>
    </ul>
  </li>
  <li><strong>Export to Excel (Comprehensive Report):</strong> Generate a detailed Excel report (.xlsx) that includes:
    <ul>
      <li><strong>Separate Sheet per Property:</strong> Each property gets its own dedicated sheet in the Excel workbook.</li>
      <li><strong>Full Details per Property:</strong> Within each property's sheet, you'll find:
        <ul>
          <li>A summary of all input fields and general property cost calculations.</li>
          <li>All three loan scenarios, each with its:
            <ul>
              <li>Detailed amortization table</li>
              <li>Rent comparison string</li>
              <li>Embedded amortization plot</li>
            </ul>
          </li>
        </ul>
      </li>
    </ul>
  </li>
</ul>

<h2>🧰 Technologies Used</h2>

<table>
  <thead>
    <tr>
      <th>Technology</th>
      <th>Purpose</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>Python 3.x</td>
      <td>Core programming language</td>
    </tr>
    <tr>
      <td>Tkinter</td>
      <td>Graphical user interface</td>
    </tr>
    <tr>
      <td>Pandas</td>
      <td>Data manipulation and calculations</td>
    </tr>
    <tr>
      <td>Matplotlib</td>
      <td>Generating amortization plots</td>
    </tr>
    <tr>
      <td>Pillow (PIL Fork)</td>
      <td>Embedding plots as images in Excel</td>
    </tr>
    <tr>
      <td>Openpyxl</td>
      <td>Writing to .xlsx files and embedding images</td>
    </tr>
  </tbody>
</table>

<h2>🚀 Getting Started</h2>

<ol>
  <li>Clone the repository</li>
  <li>Install required packages using <code>pip install -r requirements.txt</code></li>
  <li>Run the application with <code>python main.py</code></li>
</ol>

<h2>📄 License</h2>
<p>This project is licensed under the MIT License - see the <code>LICENSE</code> file for details.</p>

<h2>🙌 Contributions</h2>
<p>Pull requests are welcome! For major changes, please open an issue first to discuss what you would like to change.</p>

<img width="1039" height="901" alt="image" src="https://github.com/user-attachments/assets/8f9d1a9d-3325-4fa4-813f-8bfc08e373e4" />


<img width="968" height="843" alt="image" src="https://github.com/user-attachments/assets/38634176-52a0-4fca-9544-d8eebb1f0a2b" />


<img width="836" height="1030" alt="image" src="https://github.com/user-attachments/assets/6ac25a03-b453-4258-a763-24834063578f" />
