public final class ExcelDocumentConverter {
public static XSSFWorkbook convertWorkbookHSSFToXSSF(HSSFWorkbook source) {
    XSSFWorkbook retVal = new XSSFWorkbook();
    for (int i = 0; i < source.getNumberOfSheets(); i++) {
        XSSFSheet xssfSheet = retVal.createSheet();
        HSSFSheet hssfsheet = source.getSheetAt(i);
        copySheets(hssfsheet, xssfSheet);
    }
    return retVal;
}

public static void copySheets(HSSFSheet source, XSSFSheet destination) {
    copySheets(source, destination, true);
}

/**
 * @param destination
 *            the sheet to create from the copy.
 * @param the
 *            sheet to copy.
 * @param copyStyle
 *            true copy the style.
 */
public static void copySheets(HSSFSheet source, XSSFSheet destination, boolean copyStyle) {
    int maxColumnNum = 0;
    Map<Integer, HSSFCellStyle> styleMap = (copyStyle) ? new HashMap<Integer, HSSFCellStyle>() : null;
    for (int i = source.getFirstRowNum(); i <= source.getLastRowNum(); i++) {
        HSSFRow srcRow = source.getRow(i);
        XSSFRow destRow = destination.createRow(i);
        if (srcRow != null) {
            copyRow(source, destination, srcRow, destRow, styleMap);
            if (srcRow.getLastCellNum() > maxColumnNum) {
                maxColumnNum = srcRow.getLastCellNum();
            }
        }
    }
    for (int i = 0; i <= maxColumnNum; i++) {
        destination.setColumnWidth(i, source.getColumnWidth(i));
    }
}

/**
 * @param srcSheet
 *            the sheet to copy.
 * @param destSheet
 *            the sheet to create.
 * @param srcRow
 *            the row to copy.
 * @param destRow
 *            the row to create.
 * @param styleMap
 *            -
 */
public static void copyRow(HSSFSheet srcSheet, XSSFSheet destSheet, HSSFRow srcRow, XSSFRow destRow,
        Map<Integer, HSSFCellStyle> styleMap) {
    // manage a list of merged zone in order to not insert two times a
    // merged zone
    Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();
    destRow.setHeight(srcRow.getHeight());
    // pour chaque row
    for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {
        HSSFCell oldCell = srcRow.getCell(j); // ancienne cell
        XSSFCell newCell = destRow.getCell(j); // new cell
        if (oldCell != null) {
            if (newCell == null) {
                newCell = destRow.createCell(j);
            }
            // copy chaque cell
            copyCell(oldCell, newCell, styleMap);
            // copy les informations de fusion entre les cellules
            // System.out.println("row num: " + srcRow.getRowNum() +
            // " , col: " + (short)oldCell.getColumnIndex());
            CellRangeAddress mergedRegion = getMergedRegion(srcSheet, srcRow.getRowNum(),
                    (short) oldCell.getColumnIndex());

            if (mergedRegion != null) {
                // System.out.println("Selected merged region: " +
                // mergedRegion.toString());
                CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow(),
                        mergedRegion.getLastRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());
                // System.out.println("New merged region: " +
                // newMergedRegion.toString());
                CellRangeAddressWrapper wrapper = new CellRangeAddressWrapper(newMergedRegion);
                if (isNewMergedRegion(wrapper, mergedRegions)) {
                    mergedRegions.add(wrapper);
                    destSheet.addMergedRegion(wrapper.range);
                }
            }
        }
    }

}

/**
 * @param oldCell
 * @param newCell
 * @param styleMap
 */
public static void copyCell(HSSFCell oldCell, XSSFCell newCell, Map<Integer, HSSFCellStyle> styleMap) {
    if (styleMap != null) {
        int stHashCode = oldCell.getCellStyle().hashCode();
        HSSFCellStyle sourceCellStyle = styleMap.get(stHashCode);
        XSSFCellStyle destnCellStyle = newCell.getCellStyle();
        if (sourceCellStyle == null) {
            sourceCellStyle = oldCell.getSheet().getWorkbook().createCellStyle();
        }
        destnCellStyle.cloneStyleFrom(oldCell.getCellStyle());
        styleMap.put(stHashCode, sourceCellStyle);
        newCell.setCellStyle(destnCellStyle);
    }
    switch (oldCell.getCellType()) {
    case HSSFCell.CELL_TYPE_STRING:
        newCell.setCellValue(oldCell.getStringCellValue());
        break;
    case HSSFCell.CELL_TYPE_NUMERIC:
        newCell.setCellValue(oldCell.getNumericCellValue());
        break;
    case HSSFCell.CELL_TYPE_BLANK:
        newCell.setCellType(HSSFCell.CELL_TYPE_BLANK);
        break;
    case HSSFCell.CELL_TYPE_BOOLEAN:
        newCell.setCellValue(oldCell.getBooleanCellValue());
        break;
    case HSSFCell.CELL_TYPE_ERROR:
        newCell.setCellErrorValue(oldCell.getErrorCellValue());
        break;
    case HSSFCell.CELL_TYPE_FORMULA:
        newCell.setCellFormula(oldCell.getCellFormula());
        break;
    default:
        break;
    }

}

/**
 * Récupère les informations de fusion des cellules dans la sheet source
 * pour les appliquer à la sheet destination... Récupère toutes les zones
 * merged dans la sheet source et regarde pour chacune d'elle si elle se
 * trouve dans la current row que nous traitons. Si oui, retourne l'objet
 * CellRangeAddress.
 * 
 * @param sheet
 *            the sheet containing the data.
 * @param rowNum
 *            the num of the row to copy.
 * @param cellNum
 *            the num of the cell to copy.
 * @return the CellRangeAddress created.
 */
public static CellRangeAddress getMergedRegion(HSSFSheet sheet, int rowNum, short cellNum) {
    for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
        CellRangeAddress merged = sheet.getMergedRegion(i);
        if (merged.isInRange(rowNum, cellNum)) {
            return merged;
        }
    }
    return null;
}

/**
 * Check that the merged region has been created in the destination sheet.
 * 
 * @param newMergedRegion
 *            the merged region to copy or not in the destination sheet.
 * @param mergedRegions
 *            the list containing all the merged region.
 * @return true if the merged region is already in the list or not.
 */
private static boolean isNewMergedRegion(CellRangeAddressWrapper newMergedRegion,
        Set<CellRangeAddressWrapper> mergedRegions) {
    return !mergedRegions.contains(newMergedRegion);
}
