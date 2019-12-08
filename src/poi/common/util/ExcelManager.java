package poi.common.util;

import static poi.common.constant.ExcelConstants.*;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xddf.usermodel.chart.XDDFChart;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTCatAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLegend;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTScaling;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.STAxPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STBarDir;
import org.openxmlformats.schemas.drawingml.x2006.chart.STLegendPos;
import org.openxmlformats.schemas.drawingml.x2006.chart.STOrientation;
import org.openxmlformats.schemas.drawingml.x2006.chart.STTickLblPos;

public class ExcelManager {

	private Sheet sheet = null;
	private Row row = null;
	private Cell cell = null;
	private int offset;
	private int rowIndex;
	private int columnIndex;
	private Map<Integer,CellStyle> styleMap;

	// 設定処理

	/**
	 * シートを設定し、参照するセル位置を「A1」に設定する
	 * @param sheet
	 */
	public void setSheet(Sheet sheet) {
		this.sheet = sheet;
		setPosition();
		this.styleMap = new HashMap<>();
		this.offset = 1;
		this.rowIndex = 0;
		this.columnIndex = 0;
	}

	/**
	 * テンプレートの印刷設定をコピーする
	 * @param printSetup
	 */
	public void setPrintSetup(PrintSetup printSetup) {
		PrintSetup newSetup = this.sheet.getPrintSetup();

		// コピー枚数
		newSetup.setCopies(printSetup.getCopies());
		// 下書きモード
		//newSetup.setDraft(printSetup.getDraft());
		// シートに収まる高さのページ数
		newSetup.setFitHeight(printSetup.getFitHeight());
		// シートが収まる幅のページ数
		newSetup.setFitWidth(printSetup.getFitWidth());
		// フッター余白
		//newSetup.setFooterMargin(printSetup.getFooterMargin());
		// ヘッダー余白
		//newSetup.setHeaderMargin(printSetup.getHeaderMargin());
		// 水平解像度
		//newSetup.setHResolution(printSetup.getHResolution());
		// 横向きモード
		newSetup.setLandscape(printSetup.getLandscape());
		// 左から右への印刷順序
		//newSetup.setLeftToRight(printSetup.getLeftToRight());
		// 白黒
		newSetup.setNoColor(printSetup.getNoColor());
		// 向き
		newSetup.setNoOrientation(printSetup.getNoOrientation());
		// 印刷メモ
		//newSetup.setNotes(printSetup.getNotes());
		// ページの開始
		//newSetup.setPageStart(printSetup.getPageStart());
		// 用紙サイズ
		newSetup.setPaperSize(printSetup.getPaperSize());
		// スケール
		newSetup.setScale(printSetup.getScale());
		// 使用ページ番号
		//newSetup.setUsePage(printSetup.getUsePage());
		// 有効な設定
		//newSetup.setValidSettings(printSetup.getValidSettings());
		// 垂直解像度
		//newSetup.setVResolution(printSetup.getVResolution());

	}

	/**
	 * 印刷範囲取得
	 * @return
	 */
	public String getPrintArea() {
		int firstRow = this.sheet.getFirstRowNum();
		int lastRow = this.rowIndex;
		int firstColumn = this.sheet.getRow(firstRow).getFirstCellNum();
		int lastColumn = this.sheet.getRow(lastRow).getLastCellNum()-1;
		String printArea = "$" + getColumnAlphabet(firstColumn)
							+ "$" + String.valueOf(firstRow+1)
							+ ":$" + getColumnAlphabet(lastColumn)
							+ "$" + String.valueOf(lastRow);
		return printArea;
	}

	// 設定処理はここまで

	// 移動処理

	/**
	 * 参照する行位置を設定する
	 */
	private void setRow() {
		if((this.row = sheet.getRow(this.rowIndex)) == null) {
			this.row = sheet.createRow(this.rowIndex);
		}
	}

	/**
	 * 参照するセル位置を設定する
	 */
	private void setCell() {
		if((this.cell = row.getCell(this.columnIndex)) == null) {
			this.cell = row.createCell(this.columnIndex);
		}
	}

	/**
	 * オフセット（移動量）を設定する
	 * ※デフォルトは「1」
	 * @param offset
	 */
	public void setOffset(int offset) {
		this.offset = offset;
	}

	/**
	 * 参照する行位置、セル位置を設定する
	 */
	private void setPosition() {
		setRow();
		setCell();
	}

	/**
	 * 行と列を指定して参照するセル位置を設定する
	 * @param rowIndex
	 *          行位置(0ベース)
	 * @param columnIndex
	 *          列位置(0ベース)
	 */
	public void setPosition(int rowIndex, int columnIndex) {
		this.rowIndex = rowIndex;
		this.columnIndex = columnIndex;
		setPosition();
	}

	/**
	 * 設定してあるオフセットに従って参照するセル位置を移動する
	 * 現在列位置+オフセット
	 */
	public void nextCell() {
		moveCell(this.offset);
	}

	/**
	 * 指定したオフセットに従って参照するセル位置を移動する
	 * 現在列位置+オフセット
	 * @param columnOffset
	 */
	public void moveCell(int columnOffset) {
		move(0, columnOffset);
	}

	/**
	 * 設定してあるオフセットに従って参照する行位置を移動する
	 * 現在行位置+オフセット
	 * ※列位置は0（A）列になる
	 */
	public void nextRow() {
		nextRow(0);
	}

	/**
	 * 設定してあるオフセットに従って参照する行位置を移動し、指定した列位置に移動する
	 * 現在行位置+オフセット
	 * @param columnIndex
	 */
	public void nextRow(int columnIndex) {
		this.columnIndex = columnIndex;
		moveRow(this.offset);
	}

	/**
	 * 指定したオフセットに従って参照する行位置を移動する
	 * 現在行位置+オフセット
	 * ※列位置は変わらない
	 * @param rowOffset
	 */
	public void moveRow(int rowOffset) {
		move(rowOffset, 0);
	}

	/**
	 * 指定した行オフセット、列オフセットに従って参照する行位置、列位置を移動する
	 * 現在行位置+行オフセット
	 * 現在列位置+列オフセット
	 * @param rowOffset
	 * @param columnOffset
	 */
	public void move(int rowOffset, int columnOffset) {
		this.rowIndex += rowOffset;
		this.columnIndex += columnOffset;
		setPosition();
	}

	/**
	 * 現在行位置から以降の行を1行下へシフトする（行の挿入）
	 * ※上の行のスタイル、高さを引き継ぐ
	 */
	public void shiftRows() {
		int lastRowNum = this.sheet.getLastRowNum();
		int lastCellNum = this.sheet.getRow(this.rowIndex-1).getLastCellNum();

		this.sheet.shiftRows(this.rowIndex, lastRowNum+1, 1);
		Row newRow = this.sheet.getRow(this.rowIndex);
		if(newRow == null) {
			newRow = this.sheet.createRow(this.rowIndex);
		}
		Row oldRow = this.sheet.getRow(this.rowIndex-1);
		for(int i = 0; i < lastCellNum-1; i++) {
			Cell newCell = newRow.createCell(i);
			Cell oldCell = oldRow.getCell(i);
			// oldCellがnullでなければスタイル設定
			// wrokbookに作成出来るcellstyleには上限があるのでmapに詰める
			if(oldCell != null) {
				if(!styleMap.containsKey(i)) {
					CellStyle newStyle = this.sheet.getWorkbook().createCellStyle();
					newStyle.cloneStyleFrom(oldCell.getCellStyle());
					newStyle.setBorderTop(BorderStyle.DOTTED);
					styleMap.put(i, newStyle);
				}
				newCell.setCellStyle(styleMap.get(i));
			}
		}
		newRow.setHeightInPoints(oldRow.getHeightInPoints());
		setPosition();
	}

	// 移動処理はここまで

	// 入力処理

	/**
	 * char型の値をセルに設定し、指定したオフセットに従って参照する列位置を移動する
	 * 値がなければBLANKを設定する
	 * @param value
	 * @param columnOffset
	 */
	public void setCellValue(char value, int columnOffset) {
		setCellValue(String.valueOf(value), columnOffset);
	}
	/**
	 * char型の値をセルに設定し、設定してあるオフセットに従って参照する列位置を移動する
	 * 値がなければBLANKを設定する
	 * @param value
	 */
	public void setCellValue(char value) {
		setCellValue(value, this.offset);
	}

	/**
	 * String型の値をセルに設定し、指定したオフセットに従って参照する列位置を移動する
	 * 値がなければBLANKを設定する
	 * @param value
	 * @param columnOffset
	 */
	public void setCellValue(String value, int columnOffset) {
		if(value.trim().length() == 0) {
			this.cell.setBlank();
		}else {
			this.cell.setCellValue(value);
		}
		moveCell(columnOffset);
	}

	/**
	 * String型の値をセルに設定し、設定してあるオフセットに従って参照する列位置を移動する
	 * 値がなければBLANKを設定する
	 * @param value
	 */
	public void setCellValue(String value) {
		setCellValue(value, this.offset);
	}

	/**
	 * RichTextString型の値をセルに設定し、指定したオフセットに従って参照する列位置を移動する
	 * 値がなければBLANKを設定する
	 * @param value
	 * @param columnOffset
	 */
	public void setCellValue(RichTextString value, int columnOffset) {
		if(value.getString().trim().length() == 0) {
			this.cell.setBlank();
		}else {
			this.cell.setCellValue(value);
		}
		moveCell(columnOffset);
	}

	/**
	 * RichTextString型の値をセルに設定し、設定してあるオフセットに従って参照する列位置を移動する
	 * 値がなければBLANKを設定する
	 * @param value
	 */
	public void setCellValue(RichTextString value) {
		setCellValue(value, this.offset);
	}

	/**
	 * byte型の値をセルに設定し、指定したオフセットに従って参照する列位置を移動する
	 * @param value
	 * @param columnOffset
	 */
	public void setCellValue(byte value, int columnOffset) {
		this.cell.setCellValue(value);
		moveCell(columnOffset);
	}

	/**
	 * byte型の値をセルに設定し、設定してあるオフセットに従って参照する列位置を移動する
	 * @param value
	 */
	public void setCellValue(byte value) {
		setCellValue(value, this.offset);
	}

	/**
	 * short型の値をセルに設定し、指定したオフセットに従って参照する列位置を移動する
	 * @param value
	 * @param columnOffset
	 */
	public void setCellValue(short value, int columnOffset) {
		this.cell.setCellValue(value);
		moveCell(columnOffset);
	}

	/**
	 * short型の値をセルに設定し、設定してあるオフセットに従って参照する列位置を移動する
	 * @param value
	 */
	public void setCellValue(short value) {
		setCellValue(value, this.offset);
	}

	/**
	 * int型の値をセルに設定し、指定したオフセットに従って参照する列位置を移動する
	 * @param value
	 * @param columnOffset
	 */
	public void setCellValue(int value, int columnOffset) {
		this.cell.setCellValue(value);
		moveCell(columnOffset);
	}

	/**
	 * int型の値をセルに設定し、設定してあるオフセットに従って参照する列位置を移動する
	 * @param value
	 */
	public void setCellValue(int value) {
		setCellValue(value, this.offset);
	}

	/**
	 * long型の値をセルに設定し、指定したオフセットに従って参照する列位置を移動する
	 * @param value
	 * @param columnOffset
	 */
	public void setCellValue(long value, int columnOffset) {
		this.cell.setCellValue(value);
		moveCell(columnOffset);
	}

	/**
	 * long型の値をセルに設定し、設定してあるオフセットに従って参照する列位置を移動する
	 * @param value
	 */
	public void setCellValue(long value) {
		setCellValue(value, this.offset);
	}

	/**
	 * double型の値をセルに設定し、指定したオフセットに従って参照する列位置を移動する
	 * @param value
	 * @param columnOffset
	 */
	public void setCellValue(double value, int columnOffset) {
		this.cell.setCellValue(value);
		moveCell(columnOffset);
	}

	/**
	 * double型の値をセルに設定し、設定してあるオフセットに従って参照する列位置を移動する
	 * @param value
	 */
	public void setCellValue(double value) {
		setCellValue(value, this.offset);
	}

	/**
	 * float型の値をセルに設定し、指定したオフセットに従って参照する列位置を移動する
	 * @param value
	 * @param columnOffset
	 */
	public void setCellValue(float value, int columnOffset) {
		this.cell.setCellValue(value);
		moveCell(columnOffset);
	}

	/**
	 * float型の値をセルに設定し、設定してあるオフセットに従って参照する列位置を移動する
	 * @param value
	 */
	public void setCellValue(float value) {
		setCellValue(value, this.offset);
	}

	/**
	 * boolean型の値をセルに設定し、指定したオフセットに従って参照する列位置を移動する
	 * @param value
	 * @param columnOffset
	 */
	public void setCellValue(boolean value, int columnOffset) {
		this.cell.setCellValue(value);
		moveCell(columnOffset);
	}

	/**
	 * boolean型の値をセルに設定し、設定してあるオフセットに従って参照する列位置を移動する
	 * @param value
	 */
	public void setCellValue(boolean value) {
		setCellValue(value, this.offset);
	}

	/**
	 * Calendar型の値をセルに設定し、指定したオフセットに従って参照する列位置を移動する
	 * @param value
	 * @param columnOffset
	 */
	public void setCellValue(Calendar value, int columnOffset) {
		this.cell.setCellValue(value);
		moveCell(columnOffset);
	}

	/**
	 * Calendar型の値をセルに設定し、設定してあるオフセットに従って参照する列位置を移動する
	 * @param value
	 */
	public void setCellValue(Calendar value) {
		setCellValue(value, this.offset);
	}

	/**
	 * Date型の値をセルに設定し、指定したオフセットに従って参照する列位置を移動する
	 * @param value
	 * @param columnOffset
	 */
	public void setCellValue(Date value, int columnOffset) {
		this.cell.setCellValue(value);
		moveCell(columnOffset);
	}

	/**
	 * Date型の値をセルに設定し、設定してあるオフセットに従って参照する列位置を移動する
	 * @param value
	 */
	public void setCellValue(Date value) {
		setCellValue(value, this.offset);
	}

	/**
	 * LocalDate型の値をセルに設定し、指定したオフセットに従って参照する列位置を移動する
	 * @param value
	 * @param columnOffset
	 */
	public void setCellValue(LocalDate value, int columnOffset) {
		this.cell.setCellValue(value);
		moveCell(columnOffset);
	}

	/**
	 * LocalDate型の値をセルに設定し、設定してあるオフセットに従って参照する列位置を移動する
	 * @param value
	 */
	public void setCellValue(LocalDate value) {
		setCellValue(value, this.offset);
	}

	/**
	 * LocalDateTime型の値をセルに設定し、指定したオフセットに従って参照する列位置を移動する
	 * @param value
	 * @param columnOffset
	 */
	public void setCellValue(LocalDateTime value, int columnOffset) {
		this.cell.setCellValue(value);
		moveCell(columnOffset);
	}

	/**
	 * LocalDateTime型の値をセルに設定し、設定してあるオフセットに従って参照する列位置を移動する
	 * @param value
	 */
	public void setCellValue(LocalDateTime value) {
		setCellValue(value, this.offset);
	}

	/**
	 * String型の計算式("="は要らない)をセルに設定し、指定したオフセットに従って参照する列位置を移動する
	 * @param value
	 * @param columnOffset
	 */
	public void setCellFormula(String value, int columnOffset) {
		this.cell.setCellFormula(value);
		moveCell(columnOffset);
	}

	/**
	 * String型の計算式("="は要らない)をセルに設定し、設定してあるオフセットに従って参照する列位置を移動する
	 * @param value
	 */
	public void setCellFormula(String value) {
		setCellFormula(value, this.offset);
	}

	/**
	 * String型のコメントをセルに設定する(表示領域は固定値)
	 * ※列位置は移動しない
	 * @param commentStr
	 */
	public void setCellComment(String commentStr) {
		CreationHelper helper = sheet.getWorkbook().getCreationHelper();
		XSSFDrawing drawing = (XSSFDrawing)sheet.createDrawingPatriarch();
		// row1の位置からずらす値(ピクセル単位)
		int dx1 = 0;
		// col1の位置からずらす値(ピクセル単位)
		int dy1 = 0;
		// row2の位置からずらす値(ピクセル単位)
		int dx2 = 0;
		// col2の位置からずらす値(ピクセル単位)
		int dy2 = 0;
		// コメント表示領域の左上の列位置(セル単位)
		int col1 = this.columnIndex + 1;
		// コメント表示領域の左上の行位置(セル単位)
		int row1 = this.rowIndex;
		// コメント表示領域の右下の列位置(セル単位)
		int col2 = this.columnIndex + 4;
		// コメント表示領域の右下の行位置(セル単位)
		int row2 = this.rowIndex + 3;
		ClientAnchor anchor = drawing.createAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);
		Comment comment = drawing.createCellComment(anchor);
		//comment.setAuthor("master");
		comment.setString(helper.createRichTextString(commentStr));
		this.cell.setCellComment(comment);
	}

	/**
	 * 現在のセル位置から棒グラフを挿入(大きさは固定値)
	 */
	public void setBarChart() {
		XSSFDrawing drawing = (XSSFDrawing)sheet.createDrawingPatriarch();
		int dx1 = 0;
		int dy1 = 0;
		int dx2 = 0;
		int dy2 = 0;
		int col1 = this.columnIndex;
		int row1 = this.rowIndex;
		move(16, 10);
		int col2 = this.columnIndex;
		int row2 = this.rowIndex;
		ClientAnchor anchor = drawing.createAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);
		XDDFChart chart = (XDDFChart)drawing.createChart(anchor);

        CTChart ctChart = ((XSSFChart)chart).getCTChart();
        CTPlotArea ctPlotArea = ctChart.getPlotArea();
        CTBarChart ctBarChart = ctPlotArea.addNewBarChart();
        CTBoolean ctBoolean = ctBarChart.addNewVaryColors();
        ctBoolean.setVal(true);
        ctBarChart.addNewBarDir().setVal(STBarDir.COL);

        for (int r = 4; r < 8; r++) {
           CTBarSer ctBarSer = ctBarChart.addNewSer();
           CTSerTx ctSerTx = ctBarSer.addNewTx();
           CTStrRef ctStrRef = ctSerTx.addNewStrRef();
           // 凡例
           ctStrRef.setF("集計!$B$" + r);
           ctBarSer.addNewIdx().setVal(r-4);
           CTAxDataSource cttAxDataSource = ctBarSer.addNewCat();
           ctStrRef = cttAxDataSource.addNewStrRef();
           // 項目
           ctStrRef.setF("集計!$C$3:$I$3");
           CTNumDataSource ctNumDataSource = ctBarSer.addNewVal();
           CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
           // データエリア
           ctNumRef.setF("集計!$C$" + r + ":$I$" + r);

           ctBarSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[] {0,0,0});

        }

        int catId = 100;
        int valId = 200;
        ctBarChart.addNewAxId().setVal(catId);
        ctBarChart.addNewAxId().setVal(valId);

        CTCatAx ctCatAx = ctPlotArea.addNewCatAx();
        ctCatAx.addNewAxId().setVal(catId);
        CTScaling ctScaling = ctCatAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctCatAx.addNewDelete().setVal(false);
        ctCatAx.addNewAxPos().setVal(STAxPos.B);
        ctCatAx.addNewCrossAx().setVal(valId);
        ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

        CTValAx ctValAx = ctPlotArea.addNewValAx();
        ctValAx.addNewAxId().setVal(valId);
        ctScaling = ctValAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctValAx.addNewDelete().setVal(false);
        ctValAx.addNewAxPos().setVal(STAxPos.L);
        ctValAx.addNewCrossAx().setVal(catId);
        ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

        CTLegend ctLegend = ctChart.addNewLegend();
        ctLegend.addNewLegendPos().setVal(STLegendPos.B);
        ctLegend.addNewOverlay().setVal(false);
	}

	// 入力処理はここまで

	// 変換処理

	/**
	 * 指定した列番号をA1表記に変換する
	 * @param ascii
	 * @return
	 */
	public String getColumnAlphabet(int ascii) {
		String alphabet = "";
		if(ascii < 26) {
			alphabet = String.valueOf((char)(ASCII+ascii));
		}else {
			int div = ascii / 26;
			int mod = ascii % 26;

			alphabet = getColumnAlphabet(div-1) + getColumnAlphabet(mod);
		}
		return alphabet;
	}

	/**
	 * 現在の列番号をA1表記に変更する
	 * @return
	 */
	public String getColumnAlphabet() {
		return getColumnAlphabet(this.columnIndex);
	}

	// 変換処理はここまで

}
