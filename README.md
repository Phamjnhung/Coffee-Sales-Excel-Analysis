# Coffee-Order-Exel-Project
This project is a comprehensive data analysis solution that transforms 1,000 rows of raw coffee bean transactions into a dynamic, interactive business intelligence tool. It covers the entire data lifecycle: from complex data gathering and cleaning to advanced modeling and professional visualization.
## 📊 About Dataset

The dataset used in this project consists of 1,000 rows raw transactional data for a coffee bean retailer covering sales performance across the US, UK, and Ireland. It includes three primary tables that were integrated to build the final analysis:
* **Orders Table:** Contains transaction details including `Order ID`, `Order Date`, `Customer ID`, `Product ID`, and `Quantity`.
* **Customers Table:** Provides detailed profiles with `Customer ID`, `Customer Name`,	`Email`,	`Phone Number`,	`Address Line 1`,	`City`,	`Country`,	`Postcode`, and `Loyalty Card`.
* **Products Table:*** Includes product attributes including `Product ID`, `Coffee Type`,	`Roast Type`,	`Size`,	`Unit Price`,	`Price per 100g`,	and `Profit`. 

## 📈 Interactive Dashboard  
This project include a interactive Dashboard for exploring data.  
You can preview it right below or click into the file that I already post for further information.  

(Note: Raw data and calculation sheets are hidden to provide a user-friendly experience. Fell free to right-click any sheet tab and select 'Unhide' to view the full dataset.)

<img width="1161" height="614" alt="image" src="https://github.com/user-attachments/assets/bdde748f-e093-4efb-a828-09156436a629" />

## 🎯 Business Goals & Problem Solved
The objective was to provide the sales team with a centralized tool to answer critical business questions:

* **Trend Analysis:** Identify sales peaks and seasonal fluctuations.
* **Geographic Insights:** Compare performance across the US, UK, and Ireland.
* **Customer Profiling:** Identify top-tier customers and the impact of loyalty programs.
* **Product Strategy:** Determine which roast types and package sizes drive the most revenue.

## ⚙️ Technical Stack Summary
* **Excel Functions:** XLOOKUP, INDEX MATCH (Nested), IF logic, IFERROR.
* **Data Management:** Power Table, Data Validation, Custom Number Formatting.
* **Business Intelligence:** Pivot Tables, Pivot Charts, Slicers, Timelines, Dashboarding.

## 🛠️ Step by Step Implemetation

Step 1: Data Gathering & Transformation
The raw data was distributed across three tables: Orders, Customers, and Products. I consolidated this information into the main Orders sheet using advanced lookup techniques:

XLOOKUP for Efficient Data Retrieval: Utilized XLOOKUP to fetch Customer Names, Emails, and Countries.

Nested IF + XLOOKUP (Data Cleaning): Handled missing data by nesting XLOOKUP inside an IF statement to ensure that null values (zeros) were displayed as blanks for a cleaner dataset.

Formula: $ \text{IF}(XLOOKUP(...) = 0, "", XLOOKUP(...)) $

Dynamic INDEX & MATCH (Two-way Lookup): Instead of writing multiple lookups for each product attribute, I implemented a robust INDEX and MATCH combination. By using a nested MATCH for the column headers, I created a single dynamic formula that retrieves the correct value (Coffee Type, Roast, Size, Price) regardless of column position.

Demonstrated Skill: This approach ensures the workbook is scalable; if a new column is added to the source, the formula adapts automatically.

Step 2: Feature Engineering & Data Formatting
To enhance the analytical value of the data, I created several calculated fields and applied professional formatting:

Logical Mapping (Nested IFs): Converted raw codes (e.g., "ARA", "ROB") into descriptive names ("Arabica", "Robusta") using IF logic to make the dashboard user-friendly for non-technical stakeholders.

Custom Number Formatting: * Applied custom formatting to the weight column (e.g., 0.0 "kg") so the values remain numeric for calculations while displaying the unit.

Standardized the Date column to DD-MMM-YYYY format to avoid regional confusion (US vs. UK dates).

Formatted sales and unit prices as USD currency.

Step 3: Pivot Tables & Advanced Analysis
I converted the range into an Excel Table (orders_table) to ensure that any new data added to the bottom is automatically captured. I then built several Pivot Tables:

Time-Series Analysis: Grouped order dates by Year and Month to track seasonality.

Top 5 Customers: Applied Value Filters to dynamically isolate the top 5 spenders, allowing the business to focus on high-value accounts.

Geographic Breakdown: Aggregated sales by country to visualize market share in the US, UK, and Ireland.

Step 4: Dashboard Interactivity & Visuals
Pivot Charts: Created Line Charts for sales trends and Bar Charts for categorical comparisons, customized with a professional color palette.

Timeline Integration: Inserted a timeline to allow users to filter the entire dashboard by specific months or years with a single click.

Slicers: Added slicers for Roast Type, Size, and Loyalty Card status to provide deep-dive capabilities.

Report Connections: Crucially, I linked all Slicers and the Timeline to every Pivot Table. This ensures that a filter applied to one chart updates the entire dashboard simultaneously.

Step 5: UI/UX Polishing (Professional Look & Feel)
To transform the spreadsheet into a professional dashboard interface:

Clean Interface: Removed gridlines, row/column headers, and the formula bar.

Precision Layout: Used Alt + Drag (Snap-to-Grid) to ensure all charts and slicers were aligned with pixel-perfect accuracy.

Custom Styling: Modified the default Excel styles to create a custom "Dark Purple" theme for the timeline and slicers, ensuring a cohesive brand identity.

## 💡 Dashboard Analysis and Reccomendation
**Total Sales Over Time (Line Chart)**

<img width="1165" height="649" alt="image" src="https://github.com/user-attachments/assets/6fd57ce8-f93c-4532-ab9a-421970c82db3" />

Trend Analysis: The business shows a fluctuating but relatively consistent performance across the four-year period from 2019 to 2022. While there isn't a single uniform upward trend for all products, specific coffee types like Arabica and Liberica experience significant revenue peaks, particularly in late 2021 and early 2022, reaching highs above $800.

Seasonality & Product Volatility: Monthly revenue for individual coffee types often stays within the $100 - $400 range, but sudden spikes—such as Arabica in September 2021—suggest high-impact seasonal promotions or successful stock-replenishment cycles for premium beans.

**Market Performance by Geography (Bar Chart)**

<img width="697" height="321" alt="image" src="https://github.com/user-attachments/assets/886c12bb-fc24-46b6-aa5c-ca89c3c76ed8" />

Market Dominance: The United States stands as the primary revenue driver, contributing a staggering $35,639 in total sales, which dwarfs the combined performance of Ireland and the UK.

Performance vs. Baseline: While the US remains the stronghold, Ireland ($6,697) and the UK ($2,799) perform significantly below the market leader's baseline. This indicates a massive opportunity for expansion in European markets where the brand presence is currently underrepresented.

**Top 5 Customers (Dynamic Leaderboard)**

<img width="700" height="322" alt="image" src="https://github.com/user-attachments/assets/8b442dc1-ba5c-4268-b997-307ad2c1c087" />

VIP Concentration: The revenue from top individual customers is highly competitive, with Allis Wilmore leading at $317, followed closely by Brenn Dundredge at $307.

Customer Loyalty Potential: The small margin between the top 5 spenders (ranging from $278 to $317) shows a healthy distribution of high-value clients rather than reliance on a single major buyer.

**Dynamic Drill-Down Capabilities**

<img width="1859" height="232" alt="image" src="https://github.com/user-attachments/assets/ce77c891-674c-4ee0-884d-f5aaf515f403" />

Interactive Filtering: The power of this dashboard lies in its Cross-Filtering capability. By connecting the Slicers to multiple Pivot Tables via Report Connections, every visual stays in sync.

Self-Service Analysis: Users can perform "What-If" scenarios by combining filters. For example, a user can select "2.5kg bags" + "Non-Loyalty Customers" to see which country buys the most bulk coffee without a membership, providing a clear path for business expansion.

**🚀 Recommendation for Better Business Performance**
Based on the complete dashboard analysis, I recommend the following strategy to maximize business performance:

* **Market Expansion in Europe:** Given the massive gap between US sales ($35,639) and UK/Ireland sales ($9,496 combined), the business should launch targeted digital marketing campaigns in the UK to increase brand awareness and bridge the revenue gap.
* **Optimizing Seasonal Inventory:** Since Arabica and Liberica show high volatility with massive peaks, the supply chain team should ensure higher stock levels during peak months (like September and January) to capitalize on historical surges and prevent stockouts.
* **VIP Retention Program:** Implement a "Platinum Tier" loyalty program specifically for top spenders like Allis Wilmore and Brenn Dundredge. Offering exclusive early access to new roast types can secure their continued high-volume purchasing.
* **Roast-Specific Promotions:** Use the Roast Type Slicer to identify which roasts (Dark, Light, Medium) are underperforming in specific markets. Offering "Trial Bundles" for these specific roasts can help diversify the product mix sold in lagging regions.
# Conclusion
This project demonstrates that Excel, when used at an advanced level, can serve as a powerful BI tool. By automating data lookups and creating an interactive UI, I’ve built a solution that turns raw logs into actionable business intelligence.
