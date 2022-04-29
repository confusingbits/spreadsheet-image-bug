import {
  Alert,
  CircularProgress,
  Dialog,
  DialogActions,
  DialogContent,
  DialogContentText,
  DialogTitle,
  Button,
} from "@mui/material";
import {
  CellReferenceMode,
  IgrExcelCoreModule,
  IgrExcelModule,
  IgrExcelXlsxModule,
  Workbook,
  WorkbookLoadOptions,
  WorksheetImage,
} from "igniteui-react-excel";
import {
  IgrSpreadsheet,
  IgrSpreadsheetModule,
} from "igniteui-react-spreadsheet";
import { useEffect, useRef, useState } from "react";
import SignatureCanvas from "react-signature-canvas";

IgrExcelCoreModule.register();
IgrExcelModule.register();
IgrExcelXlsxModule.register();

IgrSpreadsheetModule.register();

interface FileViewerExcelProps {
  url: string;
}

const FileViewerExcel = ({ url }: FileViewerExcelProps) => {
  const spreadsheet = useRef<IgrSpreadsheet | null>(null);
  const signature = useRef<SignatureCanvas>(null);
  const [isLoaded, setIsLoaded] = useState(false);
  const [loadError, setLoadError] = useState("");

  const [open, setOpen] = useState(false);

  const handleClickOpen = () => {
    setOpen(true);
  };

  const handleConfirm = () => {
    setOpen(false);
    insertSignature();
  };
  const handleClose = () => {
    setOpen(false);
  };

  useEffect(() => {
    const load = async () => {
      const file = await fetch(url);
      const buffer = await file.arrayBuffer();
      Workbook.load(
        new Uint8Array(buffer),
        new WorkbookLoadOptions(),
        (w) => {
          setIsLoaded(true);
          if (!spreadsheet?.current) return;
          spreadsheet.current.workbook = w;
        },
        (e) => {
          setIsLoaded(true);
          setLoadError(e.message);
        }
      );
    };
    load();
  }, [url, spreadsheet]);

  const insertSignature = () => {
    if (!spreadsheet?.current) {
      return;
    }

    const d = signature.current?.toDataURL();

    const wi = new WorksheetImage(d);
    const s = spreadsheet.current;
    const c1 = s.activeWorksheet.getCell(
      `R${s.activeCell.row + 1}C${s.activeCell.column + 1}`,
      CellReferenceMode.R1C1
    );
    const c2 = s.activeWorksheet.getCell(
      `R${s.activeCell.row + 2}C${s.activeCell.column + 2}`,
      CellReferenceMode.R1C1
    );
    wi.topLeftCornerCell = c1;
    wi.bottomRightCornerCell = c2;

    // need to find the shape at the current location to clear it out
    const shapes = s.activeWorksheet.shapes();
    for (let i = 0; i < shapes.count; i++) {
      const shape = shapes.item(i);

      if (
        shape.topLeftCornerCell.equals(c1) &&
        shape.bottomRightCornerCell.equals(c2)
      ) {
        shapes.remove(shape);
        break;
      }
    }

    shapes.add(wi);
  };

  const handleClearSignature = () => {
    if (!spreadsheet?.current) {
      return;
    }

    const s = spreadsheet.current;
    const c1 = s.activeWorksheet.getCell(
      `R${s.activeCell.row + 1}C${s.activeCell.column + 1}`,
      CellReferenceMode.R1C1
    );
    const c2 = s.activeWorksheet.getCell(
      `R${s.activeCell.row + 2}C${s.activeCell.column + 2}`,
      CellReferenceMode.R1C1
    );

    // need to find the shape at the current location to clear it out
    const shapes = s.activeWorksheet.shapes();
    for (let i = 0; i < shapes.count; i++) {
      const shape = shapes.item(i);
      if (
        shape.topLeftCornerCell.equals(c1) &&
        shape.bottomRightCornerCell.equals(c2)
      ) {
        shapes.removeAt(i);
        break;
      }
    }
  };

  return (
    <div
      style={{ width: "100vw", height: "100vh" }}
      data-testid="file-viewer-excel"
    >
      {!isLoaded ? (
        <CircularProgress />
      ) : loadError ? (
        <Alert severity="error">{loadError}</Alert>
      ) : (
        <>
          <Button onClick={handleClickOpen}>Add Signature</Button>
          <Button onClick={handleClearSignature}>Clear Signature</Button>
          <IgrSpreadsheet ref={spreadsheet} height="100%" width="100%" />
          <Dialog
            open={open}
            onClose={handleClose}
            aria-labelledby="signature-dialog-title"
            aria-describedby="signature-dialog-description"
          >
            <DialogTitle id="signature-dialog-title">Add Signature</DialogTitle>
            <DialogContent>
              <DialogContentText id="signature-dialog-description">
                Sign Here
              </DialogContentText>
              <SignatureCanvas
                penColor="green"
                canvasProps={{
                  width: 500,
                  height: 200,
                  style: { border: "1px dotted grey" },
                  className: "sigCanvas",
                }}
                ref={signature}
              />
            </DialogContent>
            <DialogActions>
              <Button onClick={handleClose}>Cancel</Button>
              <Button onClick={handleConfirm} autoFocus>
                Confirm
              </Button>
            </DialogActions>
          </Dialog>
        </>
      )}
    </div>
  );
};

export default FileViewerExcel;
