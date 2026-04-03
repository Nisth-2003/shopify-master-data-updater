# Shopify Master Web UI

This version includes:
- Flask backend
- HTML template
- CSS styling
- Proper web UI
- Upload Shopify file
- Upload Master file
- Select sheet
- Download updated file

## Folder structure
```text
shopify_master_web_ui/
├── app.py
├── requirements.txt
├── README.md
├── uploads/
├── outputs/
├── templates/
│   └── index.html
└── static/
    └── style.css
```

## Install
```bash
pip install -r requirements.txt
```

## Run
```bash
python app.py
```

Then open:
```text
http://127.0.0.1:5000
```

## Fixed mapping
- Shopify `Name` -> Master `Order ID`
- Shopify `Email` -> Master `Email`
- Shopify `Phone` -> Master `Mobile no`
