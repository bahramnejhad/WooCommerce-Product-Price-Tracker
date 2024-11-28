# WooCommerce Product Price Tracker

A Python web scraper that tracks product prices from WooCommerce's online shop, with automatic Excel report generation and price history tracking.

## Features

- 🔄 Automatically scrapes all product pages
- 💰 Tracks price changes over time
- 📊 Generates formatted Excel reports
- 🎨 Professional Excel formatting with alternating row colors
- 🔍 Smart product matching and updating
- 📅 Price history tracking with timestamps
- 📝 RTL (Right-to-Left) support for Persian text

## Requirements

```
python >= 3.6
requests
beautifulsoup4
pandas
openpyxl
```

Install dependencies using:
```bash
pip install -r requirements.txt
```

## Usage

1. Clone the repository:
```bash
git clone https://github.com/yourusername/WooCommerce-Product-Price-Tracker.git
cd saterco-price-tracker
```

2. Run the script:
```bash
python price_tracker.py
```

The script will:
- Scan all pages of the shop
- Extract product information
- Update existing products' prices
- Add new products
- Generate a formatted Excel report

## Output

The script generates an Excel file (`all_products_list.xlsx`) with the following columns:
- Product Name (نام محصول)
- Price (قیمت)
- Last Update (تاریخ بروزرسانی)

Excel formatting includes:
- RTL layout for Persian text
- Frozen header row
- Alternating row colors
- Optimized column widths
- Centered headers with dark blue background
- Right-aligned content

## Configuration

The script uses default configurations for:
- Base URL: `https://yoursite.com/shop/`
- Output file: `all_products_list.xlsx`
- Request delay: 1 second between pages

## Functions

### `get_total_pages(soup)`
Extracts the total number of pages from the pagination section.

### `clean_price(price_text)`
Cleans and formats price text, handles missing prices.

### `scrape_page(url, headers)`
Scrapes a single page and extracts product information.

### `update_product_list(existing_df, new_products)`
Updates existing products and adds new ones while maintaining price history.

### `format_excel(filename)`
Applies professional formatting to the Excel output:
- Sets RTL direction
- Applies alternating row colors
- Centers and styles headers
- Adjusts column widths
- Freezes header row

### `scrape_saterco_shop()`
Main function that orchestrates the entire scraping process.

## Error Handling

The script includes comprehensive error handling for:
- Network connectivity issues
- Invalid HTML responses
- File access errors
- Data processing errors

## Limitations

- Respects server load with 1-second delay between requests
- Designed specifically for Saterco's website structure
- Requires stable internet connection
- Excel file should not be open during script execution

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit changes (`git commit -m 'Add AmazingFeature'`)
4. Push to branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

Distributed under the MIT License. See `LICENSE` for more information.

## Acknowledgments

- Built with BeautifulSoup4 for HTML parsing
- Uses Pandas for data manipulation
- Implements openpyxl for Excel formatting
- Designed for Persian/RTL content support

## Support

For support, please open an issue in the GitHub repository or contact the maintainers.
