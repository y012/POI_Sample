package poi.common.dao;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;

import poi.common.dto.CSVDto;

public class CSVInputDao {

	private File csvFile;
	private FileInputStream fr;
	private BufferedReader br;

	/**
	 * csvファイル設定
	 * @param file
	 */
	public void setFile(String file) {
		this.csvFile = new File(file);
	}

	/**
	 * ファイルオープン
	 * @throws IOException
	 */
	public void open() throws IOException{
		this.fr = new FileInputStream(this.csvFile);
		this.br = new BufferedReader(new InputStreamReader(this.fr, "SJIS"));
	}

	/**
	 * csvから1行ごとにDtoに取得
	 * @return
	 * @throws IOException
	 */
	public CSVDto read() throws IOException {
		String line = "";
		CSVDto csvDto = new CSVDto();

		if((line = this.br.readLine()) != null) {
			csvDto.convertFromStrArray(line.split(","));
			return csvDto;
		}else {
			return null;
		}
	}

	/**
	 * ファイルクローズ
	 * @throws IOException
	 */
	public void close() throws IOException{
		if(this.fr != null) {
			this.fr.close();
		}
		if(this.br != null) {
			this.br.close();
		}
	}

}