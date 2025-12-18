# Coffee-Order-Exel-Project
This project is a project that transforms 1,000 rows of raw coffee bean transactions into a interactive business intelligence tool. It is a end-to-end project turning complex raw data into an actionable business insights, enable the data-driven decision-making for a global retailer. 
## 📊 About Dataset

The dataset used in this project consists of 1,000 rows raw transactional data for a coffee bean retailer covering sales performance across the US, UK, and Ireland. It includes three primary tables that were integrated to build the final analysis:
* **Orders Table:** Contains transaction details including `Order ID`, `Order Date`, `Customer ID`, `Product ID`, and `Quantity`.
* **Customers Table:** Provides detailed profiles with `Customer ID`, `Customer Name`,	`Email`,	`Phone Number`,	`Address Line 1`,	`City`,	`Country`,	`Postcode`, and `Loyalty Card`.
* **Products Table:*** Includes product attributes including `Product ID`, `Coffee Type`,	`Roast Type`,	`Size`,	`Unit Price`,	`Price per 100g`,	and `Profit`. 

## 📈 Interactive Dashboard  
This project include a interactive Dashboard for exploring data.  
You can preview it right below or click into the file that I already post for further information.  

(Note: Raw data and calculation sheets are hidden to provide a user-friendly experience. Feel free to right-click any sheet tab and select 'Unhide' to view the full dataset.)

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

## 🛠️ Step by Step Implementation

**Step 1: Data Gathering & Transformation**

The raw data was distributed across three tables: `Orders`, `Customers`, and `Products`. I merge this information into the main `Orders` sheet using advanced lookup techniques:

* **XLOOKUP for Efficient Data Retrieval:** Utilized XLOOKUP to look for Customer Names, Emails, and Countries.
  * *For example:* `=XLOOKUP(C2,Customer!$A$2:$A$1001,Customer!$B$2:$B$1001,,0)`
* **Nested IF + XLOOKUP (Data Cleaning)**: Handled missing data by nesting XLOOKUP inside an IF statement to ensure that null values (zeros) were displayed as blanks for a cleaner dataset.
  * *For example:*
  
  `=IF(XLOOKUP(C2,Customer!$A$2:$A$1001,Customer!$C$2:$C$1001,,0)=0,"",XLOOKUP(C2,Customer!$A$2:$A$1001,Customer!$C$2:$C$1001,,0))`
* **Dynamic INDEX & MATCH:** By using a nested MATCH for the column headers, I created a single dynamic formula that retrieves the correct value (Coffee Type, Roast, Size, Price) regardless of column position.
   * *For example:*
  `=INDEX(Product!$A$2:$G$49, MATCH(Orders!$D2, Product!$A$2:$A$49, 0), MATCH(Orders!I$1, Product!$A$1:$G$1, 0))`
  
**Step 2: Feature Engineering & Data Formatting**
To enhance the analytical value of the data, I created several calculated fields and applied professional formatting:

* **Logical Mapping (Nested IFs):** Converted raw codes (e.g., "ARA", "ROB") into descriptive names ("Arabica", "Robusta") using IF logic to make the dashboard user-friendly for non-technical stakeholders can still understand.
* **Custom Number Formatting:** Applied custom formatting to the weight column (e.g., 0.0 "kg") so the values remain numeric for calculations while displaying the unit.
* **Standardized the Date column**: turn it into DD-MMM-YYYY format to avoid regional confusion (US vs. UK dates).
* **Formatted sales and unit prices as USD currency.**

**Step 3: Pivot Tables & Advanced Analysis**
I converted the range into an Excel Table to ensure that any new data added to the bottom is automatically captured. I then built several Pivot Tables:

* **Time-Series Analysis:** Grouped order dates by Year and Month to track seasonality.
* **Top 5 Customers:** Applied Value Filters to dynamically isolate the top 5 spenders, allowing the business to focus on big spender.
* **Geographic Breakdown:** Aggregated sales by country to visualize market share in the US, UK, and Ireland.

**Step 4: Dashboarding**
I create a Dashboard to easily exploring data, trend that help in business decision-making.

* **Pivot Charts:** Created Line Charts for sales trends and Bar Charts for 'Hight Ticket' country and personal buyer comparisons, customized with a professional color palette.
* **Timeline Integration:** Inserted a timeline to allow users to filter the entire dashboard by specific months or years with a single click.
* **Slicers:** Added slicers for Roast Type, Size, and Loyalty Card status to provide deep-dive capabilities.
* **Report Connections:** Crucially, I linked all Slicers and the Timeline to every Pivot Table. This ensures that a filter applied to one chart updates the entire dashboard simultaneously.

## 💡 Dashboard Analysis and Reccomendation
**Total Sales Over Time (Line Chart)**

<img width="1165" height="649" alt="image" src="https://github.com/user-attachments/assets/6fd57ce8-f93c-4532-ab9a-421970c82db3" />

The graph show that cafe sales were unstable from 2019 to 2022. Price and demand change quicky across four type of coffee, there is no specific pattern for any type.

* **The fluctuation:** Every coffee type had "peaks" and "valleys" (low sales). For example, Arabica reached its highest point in September 2021, but fell deeply several times before that. This shows the market is very unpredictable. Example: Liberica sales went from more than $800 in early 2022 to nearly $0 in August 2022. This shows how fast the market changes.
* **Customer Switching:** When one coffee type became less popular, another often went up. This suggests that if one type is too expensive or out of stock, customers simply buy a different one. Example: In early 2019, Excelsa sales were very high while Robusta was very low, showing customers picked one over the other.
* **Robusta is the Most Stable:** While other coffees had huge spikes, Robusta usually stayed at a lower level. It did not reach $800 like the others, but it also had fewer extreme decline. Example: In 2019, Robusta stayed mostly between $100 and $300, making it the most "predictable" product in the group.
* **A Sharp Drop in 2022:** At the very end of the graph (August 2022), the sales for all four types dropped significantly. This is a bad sign, showing that the whole coffee market struggled at that time.

**Market Performance by Geography (Bar Chart)**

<img width="697" height="321" alt="image" src="https://github.com/user-attachments/assets/886c12bb-fc24-46b6-aa5c-ca89c3c76ed8" />

The graph shows the sales proportion for three countries: United States, Ireland, and United Kingdom.

* **Market Dominance:** The United States stands as the primary revenue driver, contributing $35,639 in total sales, which still bigger than total sales between Ireland and United Kingdom combined.
* **Performance vs. Baseline:** While the US remains the stronghold, Ireland ($6,697) and the UK ($2,799) perform significantly below the market leader's baseline. This indicates a massive opportunity for expansion in other markets to maximize the revenue.

**Top 5 Customers (Dynamic Leaderboard)**

<img width="700" height="322" alt="image" src="https://github.com/user-attachments/assets/8b442dc1-ba5c-4268-b997-307ad2c1c087" />

The bar chart tell us the top 5 customer that a big spender that the business should focus on loyalty service.
* **Big fish** The biggest spender iss Allis Wilmore with a competitive revenue, at $317 and following by Brenn Dundredge at $307
* **Customer Loyalty Potential:** The small margin between the top 5 spenders (ranging from $278 to $317) shows a healthy distribution of high-value clients rather than just lie on one top customer.

**Dynamic Drill-Down Capabilities**

<img width="1859" height="232" alt="image" src="https://github.com/user-attachments/assets/ce77c891-674c-4ee0-884d-f5aaf515f403" />

* **Interactive Filtering:** The power of this dashboard lies in its Cross-Filtering capability. By connecting the Slicers to multiple Pivot Tables via Report Connections, every graph stays in sync.
* **Self-Service Analysis:** Users can perform combining filters to explore all scenario. For example, a user can select "2.5kg bags" + "Non-Loyalty Customers" to see which country buys the most bulk coffee without a membership, providing a clear path for business expansion.

**🚀 Recommendation for Better Business Performance**
Based on the complete dashboard analysis, I recommend the following strategy to maximize business performance:

* **Market Expansion in other country:** Given the massive gap between US sales ($35,639) and UK/Ireland sales ($9,496 combined), the business should launch digital marketing campaigns to the UK/Ireland to increase brand awareness and also explore a new entry in available country to expanding the business.
* **Optimizing Seasonal Inventory:** Since Arabica and Liberica show high volatility with massive peaks, the supply chain team should ensure higher stock levels during peak months (like September and January) to capitalize on previous surges and prevent stockouts.
* **VIP Retention Program:** Implement a ranking loyalty program specifically for top spenders like Allis Wilmore and Brenn Dundredge that offering a discout voucher or some special treatment to keep them as a loyalty customer. This also encourage buyer to purchase more to achieve a high ranking with more loyalty customer benefit.

# Conclusion
This project illustrates that Excel, when used at an advanced level, can serve as a powerful BI tool. By automating data lookups and creating an interactive visualization, I’ve built a solution that turns raw data into actionable business intelligence.

