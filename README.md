
# AI myDATA Automation — CRM (DocAI / Vision)

## Γρήγορα βήματα
```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
```

1) Βάλε **vision.json** (Service Account Key) δίπλα στο `app.py`.
2) Άνοιξε `docai_config.json`:
   - `"provider": "docai"` για Document AI ή `"vision"` για Google Vision.
   - Συμπλήρωσε `project_id`, `location`, `processor_id_invoice`/`processor_id_expense` (αν χρησιμοποιήσεις DocAI).
3) Τρέξε:
```bash
uvicorn app:app --reload
```
4) Άνοιξε `http://localhost:8000`

## Σελίδες
- `/` → Πελάτες (search, create)
- `/static/customer.html?cid=ID` → Καρτέλα πελάτη: upload τιμολόγια/αποδείξεις, inline edit, φίλτρα, confidence, γραφήματα, export Excel.

## Σημειώσεις
- Αν ανεβάζεις **PDF** και έχεις `provider=docai`, χρησιμοποιεί Document AI. Αν όχι, μετατρέπεται σε εικόνες και περνάει από Vision.
- Η στήλη **Conf.** είναι το μέσο confidence (κόκκινο αν < 80%).
- Inline edit ενημερώνει backend και κάνει auto υπολογισμούς net/vat/total.
