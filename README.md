# Dashboard Aset SAP vs eAsset

## Cara jalankan

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy ke Streamlit Community Cloud

1. Masukkan `app.py`, `requirements.txt` dan `Asset.xlsx` ke GitHub repository.
2. Buka Streamlit Community Cloud.
3. Pilih repository dan tetapkan main file kepada `app.py`.
4. Deploy.

Dashboard juga menyediakan fungsi upload Excel supaya data boleh dikemas kini tanpa mengubah kod.

## Struktur sheet diperlukan

- `SAP`
- `EAsset`
- `DIM Eva grp 1`

Padanan dibuat menggunakan medan `Asset` dalam SAP dan `No. Aset SAP` dalam eAsset.
