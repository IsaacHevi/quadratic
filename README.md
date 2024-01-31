This is a Python script that creates an Excel file containing two sheets: Sheet1 and Sheet2. The first sheet (Sheet1) is populated with 100 rows of random coefficients for quadratic equations. The second sheet (Sheet2) then stores the roots of these quadratic equations.
## FEATURES
#### Sheet1 Generation:
<p>The script generates 100 rows of random coefficients for quadratic equations in Sheet1. Each row has three columns: A, B, and C.</p>
<p>Column A contains random integers from 1 to 10.</p>
<p>Columns B and C contain random integers from 0 to 10.</p>
#### Roots Calculation:
<p>The roots of each quadratic equation in Sheet1 are calculated using the quadratic formula.</p>
#### Sheet2 Population:
<p>The roots are stored in Sheet2 of the Excel file.</p>
<p>If a root has an imaginary part, it is stored as a complex number.</p>
<p>If a root is purely real, it is stored as a real number.</p>
## REQUIREMENTS
<p>Python 3.x</p>
<p>openpyxl library (install using pip install openpyxl)</p>
