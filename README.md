# Unicommerce-API-Sheets-Automation
A Google Apps Script solution that integrates with the Unicommerce e-commerce platform.
This script automatically fetches order data, keeps it synchronized in real time with Google Sheets, manages OAuth tokens, and allows customizable data filtering.
It also includes built-in logic to detect and prevent duplicate orders along with robust error handling.

Features
OAuth password grant-based authentication to Unicommerce API.

Fetch sale orders using /saleOrder/search endpoint with customizable date filters.

Fetch detailed order information with /saleorder/get.

Data writing to Google Sheets with duplicate entry avoidance.

Configurable columns with deep field mapping from API response.

Automated trigger setup to periodically fetch data.

Comprehensive error handling and logging.

Formats date/time columns in Sheets for better readability.
