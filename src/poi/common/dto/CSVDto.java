package poi.common.dto;

public class CSVDto {

	private int ELEMENT = 12;

	// カラム名があるため型はすべてString
	// 名前
	private String name;
	// ふりがな
	private String furigana;
	// メールアドレス
	private String mail;
	// 性別
	private String sex;
	// 年齢
	private String age;
	// 誕生日
	private String birthDay;
	// 婚姻
	private String marriage;
	// 血液型
	private String bloodType;
	// 出身地（都道府県）
	private String birthPlace;
	// 電話番号
	private String telephone;
	// 携帯
	private String mobile;
	// キャリア
	private String Career;

	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getFurigana() {
		return furigana;
	}
	public void setFurigana(String furigana) {
		this.furigana = furigana;
	}
	public String getMail() {
		return mail;
	}
	public void setMail(String mail) {
		this.mail = mail;
	}
	public String getSex() {
		return sex;
	}
	public void setSex(String sex) {
		this.sex = sex;
	}
	public String getAge() {
		return age;
	}
	public void setAge(String age) {
		this.age = age;
	}
	public String getBirthDay() {
		return birthDay;
	}
	public void setBirthDay(String birthDay) {
		this.birthDay = birthDay;
	}
	public String getMarriage() {
		return marriage;
	}
	public void setMarriage(String marriage) {
		this.marriage = marriage;
	}
	public String getBloodType() {
		return bloodType;
	}
	public void setBloodType(String bloodType) {
		this.bloodType = bloodType;
	}
	public String getBirthPlace() {
		return birthPlace;
	}
	public void setBirthPlace(String birthPlace) {
		this.birthPlace = birthPlace;
	}
	public String getTelephone() {
		return telephone;
	}
	public void setTelephone(String telephone) {
		this.telephone = telephone;
	}
	public String getMobile() {
		return mobile;
	}
	public void setMobile(String mobile) {
		this.mobile = mobile;
	}
	public String getCareer() {
		return Career;
	}
	public void setCareer(String career) {
		Career = career;
	}

	/**
	 * Dto→String配列に変換する
	 * @return
	 */
	public String[] convertToStrArray() {
		String[] strArray = new String[ELEMENT];
		int count = 0;

		strArray[count++] = this.getName();
		strArray[count++] = this.getFurigana();
		strArray[count++] = this.getMail();
		strArray[count++] = this.getSex();
		strArray[count++] = this.getAge();
		strArray[count++] = this.getBirthDay();
		strArray[count++] = this.getMarriage();
		strArray[count++] = this.getBloodType();
		strArray[count++] = this.getBirthPlace();
		strArray[count++] = this.getTelephone();
		strArray[count++] = this.getMobile();
		strArray[count++] = this.getCareer();

		return strArray;
	}

	/**
	 * String配列→Dtoに変換する
	 * @param strArray
	 */
	public void convertFromStrArray(String[] strArray) {
		int count = 0;

		this.setName(strArray[count++]);
		this.setFurigana(strArray[count++]);
		this.setMail(strArray[count++]);
		this.setSex(strArray[count++]);
		this.setAge(strArray[count++]);
		this.setBirthDay(strArray[count++]);
		this.setMarriage(strArray[count++]);
		this.setBloodType(strArray[count++]);
		this.setBirthPlace(strArray[count++]);
		this.setTelephone(strArray[count++]);
		this.setMobile(strArray[count++]);
		this.setCareer(strArray[count++]);
	}
}