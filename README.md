# HKTVmall Product Scraper

Streamlit web app for scraping HKTVmall product data with unlimited coverage via automatic task splitting.

## Features

- Keyword, category, combined, and all-products scraping modes
- Automatic task splitting when the API pagination window (~10,000 offset) is reached:
  - Sub-category discovery from product `primaryCatCode`
  - Price range buckets
  - Brand facets
- Excel export with embedded product images (ZIP package)
- Checkpoint / backup / resume support
- Traditional Chinese UI

## Setup

```bash
pip install -r requirements.txt
streamlit run hktvmall_scraper.py
```

## Testing

```bash
python3 test_scraper_core.py
```

## Notes

- Set **商品上限** to `0` for unlimited scraping (default).
- Large category runs may take a long time; use checkpoints and auto-backup.
- The HKTVmall search API caps deep pagination; the scraper works around this automatically.
