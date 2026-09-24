import io
import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware

# استدعاء الملفات بالسميات الصحيحة ديالهم
import app123
import apvsbmcil
import apvscaml

app = FastAPI(title="Reconciliation API Engine")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def read_root():
    return {"status": "API Reconciliation Online"}

@app.post("/reconcile")
async def reconcile(
    engine: str = Form(...),  # "engine_1", "engine_2", ou "engine_3"
    erp_file: UploadFile = File(...),
    rlv_file: UploadFile = File(...)
):
    try:
        erp_bytes = await erp_file.read()
        rlv_bytes = await rlv_file.read()

        # ENGINE 1
        if engine == "engine_1":
            erp_raw = app123.read_excel(io.BytesIO(erp_bytes))
            rlv_raw = app123.read_excel(io.BytesIO(rlv_bytes))
            
            erp, _ = app123.normalize(erp_raw)
            rlv, _ = app123.normalize(rlv_raw)
            
            matches = app123.reconcile(erp, rlv)
            _, matched_df, un_erp, un_rlv = app123.build_report(erp_raw, rlv_raw, erp, rlv, matches)

            return {
                "status": "success",
                "matched": matched_df.fillna("").astype(str).to_dict(orient="records"),
                "unmatched_erp": un_erp.fillna("").astype(str).to_dict(orient="records"),
                "unmatched_rlv": un_rlv.fillna("").astype(str).to_dict(orient="records"),
            }

        # ENGINE 2
        elif engine == "engine_2":
            erp_df = apvsbmcil.prepare(apvsbmcil.read_upload(io.BytesIO(erp_bytes)))
            rlv_df = apvsbmcil.prepare(apvsbmcil.read_upload(io.BytesIO(rlv_bytes)))
            
            matches, em, rm = apvsbmcil.reconcile(erp_df, rlv_df)
            _, matched_df, unmatched_df = apvsbmcil.build_outputs(erp_df, rlv_df, matches, em, rm)

            un_erp = unmatched_df[unmatched_df["Source"] == "ERP"]
            un_rlv = unmatched_df[unmatched_df["Source"] == "RLV"]

            return {
                "status": "success",
                "matched": matched_df.fillna("").astype(str).to_dict(orient="records"),
                "unmatched_erp": un_erp.fillna("").astype(str).to_dict(orient="records"),
                "unmatched_rlv": un_rlv.fillna("").astype(str).to_dict(orient="records"),
            }

        # ENGINE 3 (apvscaml.py)
        elif engine == "engine_3":
            df_erp, df_rlv = apvscaml.charger_donnees(erp_bytes)
            matched_df, unmatched_rlv_df = apvscaml.lancer_rapprochement(df_erp, df_rlv)

            return {
                "status": "success",
                "matched": matched_df.fillna("").astype(str).to_dict(orient="records"),
                "unmatched_erp": [],
                "unmatched_rlv": unmatched_rlv_df.fillna("").astype(str).to_dict(orient="records"),
            }

        else:
            return {"status": "error", "message": "المحرك غير معروف"}

    except Exception as e:
        return {"status": "error", "message": str(e)}