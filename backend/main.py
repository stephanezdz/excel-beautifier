from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from typing import Optional
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from io import BytesIO
import uuid

app = FastAPI(title="Excel Beautifier API")

class BeautificationSettings(BaseModel):
    theme: str = "professional"
    header_font_size: int = 14
    header_font_weight: str = "bold"
    header_color: str = "#2E74B5"
    alternate_row_color: str = "#F2F2F2"
    border_color: str = "#CCCCCC"
    header_alignment: str = "center"
    apply_conditional_formatting: bool = True

def apply_beautification(df: pd.DataFrame, settings: BeautificationSettings) -> bytes:
    """Apply beautification rules to Excel file"""
    
    # Create workbook
    wb = load_workbook()
    ws = wb.active
    
    # Convert DataFrame to Excel
    df.to_excel(ws, index=False, header=True, start_row=1, start_col=1)
    
    # Get dimensions
    max_cols = df.shape[1]
    max_rows = df.shape[0] + 1  # +1 for header
    
    # Apply header formatting
    header_font = Font(
        name="Calibri",
        size=settings.header_font_size,
        bold=True if settings.header_font_weight == "bold" else False,
        color=settings.header_color
    )
    
    header_fill = PatternFill(
        fill_type="solid",
        color=settings.header_color
    )
    
    header_alignment = Alignment(
        horizontal=settings.header_alignment,
        vertical="center"
    )
    
    # Apply border
    thin_border = Border(
        left=Side(color=settings.border_color, style="thin"),
        right=Side(color=settings.border_color, style="thin"),
        top=Side(color=settings.border_color, style="thin"),
        bottom=Side(color=settings.border_color, style="thin")
    )
    
    # Format header row
    for col in range(1, max_cols + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        
    # Apply alternate row colors (zebra striping)
    if settings.theme in ["professional", "modern"]:
        for row in range(3, max_rows + 1):
            if row % 2 == 0:
                fill = PatternFill(fill_type="solid", color=settings.alternate_row_color)
                for col in range(1, max_cols + 1):
                    ws.cell(row=row, column=col).fill = fill
    
    # Apply borders to all cells
    for row in range(1, max_rows + 1):
        for col in range(1, max_cols + 1):
            ws.cell(row=row, column=col).border = thin_border
    
    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output.getvalue()

@app.post("/beautify")
async def beautify_excel(
    file: UploadFile = File(...),
    settings: BeautificationSettings = Form(...)
):
    """Main endpoint to beautify Excel file"""
    
    try:
        # Read Excel file
        contents = await file.read()
        df = pd.read_excel(BytesIO(contents))
        
        # Apply beautification
        beautified = apply_beautification(df, settings)
        
        # Return as response
        return JSONResponse(
            content={"status": "success", "data": beautified},
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {"status": "healthy"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
