package me.chyxion.xls;

import me.chyxion.xls.css.CssApplier;
import me.chyxion.xls.css.support.*;
import me.chyxion.xls.model.SheetDTO;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;

/**
 * @version 0.0.1
 * @since 0.0.1
 * @author Shaun Chyxion <br>
 * chyxion@163.com <br>
 * Oct 24, 2014 2:09:02 PM
 */
public class TableListToXls {
	private static final Logger log = 
		LoggerFactory.getLogger(TableListToXls.class);
	private static final List<CssApplier> STYLE_APPLIERS = 
		new LinkedList<CssApplier>();
	// static init
	static {
		STYLE_APPLIERS.add(new AlignApplier());
		STYLE_APPLIERS.add(new BackgroundApplier());
		STYLE_APPLIERS.add(new WidthApplier());
		STYLE_APPLIERS.add(new HeightApplier());
		STYLE_APPLIERS.add(new BorderApplier());
		STYLE_APPLIERS.add(new TextApplier());
	}
	private HSSFWorkbook workBook = new HSSFWorkbook();
	private Map<String, Map<String, Object>> sheetCellsOccupied = new HashMap<String, Map<String, Object>>();
	private Map<String, Map<String, HSSFCellStyle>> sheetCellStyles = new HashMap<String, Map<String, HSSFCellStyle>>();
	private int maxRow = 0;


	/**
	 * process html to xls
	 * @param sheets SheetDTO
	 * @return xls bytes
	 */
	public static byte[] process(SheetDTO[] sheets) {
		ByteArrayOutputStream baos = null;
		try {
			baos = new ByteArrayOutputStream();
			process(sheets, baos);
			return baos.toByteArray();
		}
		finally {
			if (baos != null) {
				try {
					baos.close();
				}
				catch (IOException e) {
					log.warn("Close Byte Array Inpout Stream Error Caused.", e);
				}
			}
		}
	}

	private HSSFCellStyle getDefaultCellStyle(){
		HSSFCellStyle defaultCellStyle = workBook.createCellStyle();
		defaultCellStyle.setWrapText(true);
		defaultCellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		// border
		short black = new HSSFColor.BLACK().getIndex();
		short thin = CellStyle.BORDER_THIN;
		// top
		defaultCellStyle.setBorderTop(thin);
		defaultCellStyle.setTopBorderColor(black);
		// right
		defaultCellStyle.setBorderRight(thin);
		defaultCellStyle.setRightBorderColor(black);
		// bottom
		defaultCellStyle.setBorderBottom(thin);
		defaultCellStyle.setBottomBorderColor(black);
		// left
		defaultCellStyle.setBorderLeft(thin);
		defaultCellStyle.setLeftBorderColor(black);
		return defaultCellStyle;
	}

	/**
	 * process html to output stream
	 * @param sheets html SheetDTO[]
	 * @param output output stream
	 */
	public static void process(SheetDTO[] sheets, OutputStream output) {
		new TableListToXls().doProcess(sheets, output);
	}

	// --
	// private methods

	private void processTable(HSSFSheet sheet,Element table,String sheetName) {
		int rowIndex = 0;
		if (maxRow > 0) {
			// blank row
			maxRow += 2;
			rowIndex = maxRow;
		}
		log.info("Interate Table Rows.");
		for (Element row : table.select("tr")) {
			log.info("Parse Table Row [{}]. Row Index [{}].", row, rowIndex);
			int colIndex = 0;
			log.info("Interate Cols.");
			for (Element td : row.select("td, th")) {
				// skip occupied cell
				Map<String, Object> cellsOccupied = getSheetCellsOccupied(sheetName);
				while (cellsOccupied.get(rowIndex + "_" + colIndex) != null) {
					log.info("Cell [{}][{}] Has Been Occupied, Skip.", rowIndex, colIndex);
					++colIndex;
				}
				log.info("Parse Col [{}], Col Index [{}].", td, colIndex);
				int rowSpan = 0;
				String strRowSpan = td.attr("rowspan");
				if (StringUtils.isNotBlank(strRowSpan) && 
						StringUtils.isNumeric(strRowSpan)) {
					log.info("Found Row Span [{}].", strRowSpan);
					rowSpan = Integer.parseInt(strRowSpan);
				}
				int colSpan = 0;
				String strColSpan = td.attr("colspan");
				if (StringUtils.isNotBlank(strColSpan) && 
						StringUtils.isNumeric(strColSpan)) {
					log.info("Found Col Span [{}].", strColSpan);
					colSpan = Integer.parseInt(strColSpan);
				}
				// col span & row span
				if (colSpan > 1 && rowSpan > 1) {
					spanRowAndCol(sheet,td, rowIndex, colIndex, rowSpan, colSpan,sheetName);
					colIndex += colSpan;
				}
				// col span only
				else if (colSpan > 1) {
					spanCol(sheet,td, rowIndex, colIndex, colSpan,sheetName);
					colIndex += colSpan;
				}
				// row span only
				else if (rowSpan > 1) {
					spanRow(sheet,td, rowIndex, colIndex, rowSpan,sheetName);
					++colIndex;
				}
				// no span
				else {
					createCell(td, getOrCreateRow(sheet, rowIndex), colIndex,sheetName).setCellValue(td.text());
					++colIndex;
				}
			}
			++rowIndex;
		}
	}

	private void doProcess(SheetDTO[] sheets, OutputStream output) {
		for (SheetDTO dto : sheets) {
			HSSFSheet sheet = workBook.createSheet(dto.getSheetName());
			for (Element table : Jsoup.parseBodyFragment(dto.getHtml()).select("table")) {
				processTable(sheet, table, dto.getSheetName());
			}
			maxRow = 0;
		}
		try {
			workBook.write(output);
		}
		catch (IOException e) {
			throw new IllegalStateException("Table To XLS, IO ERROR.", e);
		}
	}

    private void spanRow(HSSFSheet sheet,Element td, int rowIndex, int colIndex, int rowSpan,String sheetName) {
    	log.info("Span Row , From Row [{}], Span [{}].", rowIndex, rowSpan);
    	mergeRegion(sheet,rowIndex, rowIndex + rowSpan - 1, colIndex, colIndex);
		Map<String, Object> cellsOccupied = getSheetCellsOccupied(sheetName);
		for (int i = 0; i < rowSpan; ++i) {
    		HSSFRow row = getOrCreateRow(sheet,rowIndex + i);
    		createCell(td, row, colIndex,sheetName);
    		cellsOccupied.put((rowIndex + i) + "_" + colIndex, true);
    	}
    	getOrCreateRow(sheet,rowIndex).getCell(colIndex).setCellValue(td.text());
    }

    private void spanCol(HSSFSheet sheet,Element td, int rowIndex, int colIndex, int colSpan,String sheetName) {
    	log.info("Span Col, From Col [{}], Span [{}].", colIndex, colSpan);
    	mergeRegion(sheet,rowIndex, rowIndex, colIndex, colIndex + colSpan - 1);
    	HSSFRow row = getOrCreateRow(sheet,rowIndex);
    	for (int i = 0; i < colSpan; ++i) {
    		createCell(td, row, colIndex + i,sheetName);
    	}
    	row.getCell(colIndex).setCellValue(td.text());
    }

    private void spanRowAndCol(HSSFSheet sheet,Element td, int rowIndex, int colIndex,
            int rowSpan, int colSpan,String sheetName) {
    	log.info("Span Row And Col, From Row [{}], Span [{}].", rowIndex, rowSpan);
    	log.info("From Col [{}], Span [{}].", colIndex, colSpan);
    	mergeRegion(sheet,rowIndex, rowIndex + rowSpan - 1, colIndex, colIndex + colSpan - 1);
		Map<String, Object> cellsOccupied = getSheetCellsOccupied(sheetName);
		for (int i = 0; i < rowSpan; ++i) {
    		HSSFRow row = getOrCreateRow(sheet,rowIndex + i);
    		for (int j = 0; j < colSpan; ++j) {
    			createCell(td, row, colIndex + j,sheetName);
    			cellsOccupied.put((rowIndex + i) + "_" + (colIndex + j), true);
    		}
    	}
    	getOrCreateRow(sheet, rowIndex).getCell(colIndex).setCellValue(td.text());
    }

    private HSSFCell createCell(Element td, HSSFRow row, int colIndex,String sheetName) {
    	HSSFCell cell = row.getCell(colIndex);
    	if (cell == null) {
    		log.debug("Create Cell [{}][{}].", row.getRowNum(), colIndex);
    		cell = row.createCell(colIndex);
    	}
    	return applyStyle(td, cell,sheetName);
    }

    private HSSFCell applyStyle(Element td, HSSFCell cell,String sheetName) {
    	String style = td.attr(CssApplier.STYLE);
    	HSSFCellStyle cellStyle = null;
    	if (StringUtils.isNotBlank(style)) {
			Map<String, HSSFCellStyle> cellStyles = getSheetCellStyles(sheetName);
			log.debug("sheetName: {}, cellStyles.size(): {}",sheetName,cellStyles.size());
    		if (cellStyles.size() < 4000) {
				Map<String, String> mapStyle = parseStyle(style.trim());
				Map<String, String> mapStyleParsed = new HashMap<String, String>();
				for (CssApplier applier : STYLE_APPLIERS) {
					mapStyleParsed.putAll(applier.parse(mapStyle));
				}
				cellStyle = cellStyles.get(styleStr(mapStyleParsed));
				if (cellStyle == null) {
					log.debug("No Cell Style Found In Cache, Parse New Style.");
					cellStyle = workBook.createCellStyle();
					cellStyle.cloneStyleFrom(getDefaultCellStyle());
					for (CssApplier applier : STYLE_APPLIERS) {
						applier.apply(cell, cellStyle, mapStyleParsed);
					}
					// cache style
					cellStyles.put(styleStr(mapStyleParsed), cellStyle);
				}
    		}
    		else {
    			log.info("Custom Cell Style Exceeds 4000, Could Not Create New Style, Use Default Style.");
    			cellStyle = getDefaultCellStyle();
    		}
    	}
    	else {
    		log.debug("Use Default Cell Style.");
    		cellStyle = getDefaultCellStyle();
    	}
    	cell.setCellStyle(cellStyle);
	    return cell;
    }

	private Map<String, HSSFCellStyle> getSheetCellStyles(String sheetName){
		Map<String, HSSFCellStyle> cellStyleMap = sheetCellStyles.get(sheetName);
		if(cellStyleMap == null){
			cellStyleMap = new HashMap<String, HSSFCellStyle>();
			sheetCellStyles.put(sheetName, cellStyleMap);
		}
		return cellStyleMap;
	}

	private Map<String,Object> getSheetCellsOccupied(String sheetName){
		Map<String,Object> cellsOccupied = sheetCellsOccupied.get(sheetName);
		if(cellsOccupied == null){
			cellsOccupied = new HashMap<String, Object>();
			sheetCellsOccupied.put(sheetName,cellsOccupied);
		}
		return cellsOccupied;
	}


    private String styleStr(Map<String, String> style) {
    	log.debug("Build Style String, Style [{}].", style);
    	StringBuilder sbStyle = new StringBuilder();
    	Object[] keys = style.keySet().toArray();
    	Arrays.sort(keys);
    	for (Object key : keys) {
    		sbStyle.append(key)
    		.append(':')
    		.append(style.get(key))
    		.append(';');
        }
    	log.debug("Style String Result [{}].", sbStyle);
    	return sbStyle.toString();
    }

    private Map<String, String> parseStyle(String style) {
    	log.debug("Parse Style String [{}] To Map.", style);
    	Map<String, String> mapStyle = new HashMap<String, String>();
    	for (String s : style.split("\\s*;\\s*")) {
    		if (StringUtils.isNotBlank(s)) {
    			String[] ss = s.split("\\s*\\:\\s*");
    			if (ss.length == 2 &&
    					StringUtils.isNotBlank(ss[0]) &&
    					StringUtils.isNotBlank(ss[1])) {
    				String attrName = ss[0].toLowerCase();
    				String attrValue = ss[1];
    				// do not change font name
    				if (!CssApplier.FONT.equals(attrName) && 
    					!CssApplier.FONT_FAMILY.equals(attrName)) {
    					attrValue = attrValue.toLowerCase();
    				}
    				mapStyle.put(attrName, attrValue);
    			}
    		}
    	}
    	log.debug("Style Map Result [{}].", mapStyle);
	    return mapStyle;
    }

    private HSSFRow getOrCreateRow(HSSFSheet sheet,int rowIndex) {
    	HSSFRow row = sheet.getRow(rowIndex);
    	if (row == null) {
    		log.info("Create New Row [{}].", rowIndex);
    		row = sheet.createRow(rowIndex);
    		if (rowIndex > maxRow) {
    			maxRow = rowIndex;
    		}
    	}
	    return row;
    }

    private void mergeRegion(HSSFSheet sheet,int firstRow, int lastRow, int firstCol, int lastCol) {
    	log.debug("Merge Region, From Row [{}], To [{}].", firstRow, lastRow);
    	log.debug("From Col [{}], To [{}].", firstCol, lastCol);
    	sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }
}
