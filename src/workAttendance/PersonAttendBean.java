package workAttendance;

public class PersonAttendBean {
	String name;
	String position;
	String kqkk;
	String kqyc;
	int weishua;
	int chidao;
	int zaotui;
	int mqDays;
//	int mqWeeks;
	String hege;
	String zhiban;

	public PersonAttendBean(){
		this.weishua = 0;
		this.chidao = 0;
		this.zaotui = 0;
		this.mqDays = 0;
//		this.mqWeeks = 0;
		this.position="";
		this.kqkk="";
		this.kqyc="";
		this.hege="";
		this.zhiban="";
	}
	public PersonAttendBean(String name){
		this.name = name;
		this.weishua = 0;
		this.chidao = 0;
		this.zaotui = 0;
		this.mqDays = 0;
//		this.mqWeeks = 0;
		this.position="";
		this.kqkk="";
		this.kqyc="";
		this.hege="";
		this.zhiban="";
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getPosition() {
		return position;
	}
	public void setPosition(String position) {
		this.position = position;
	}
	public String getKqkk() {
		return kqkk;
	}
	public void setKqkk(String kqkk) {
		this.kqkk = kqkk;
	}
	public String getKqyc() {
		return kqyc;
	}
	public void setKqyc(String kqyc) {
		this.kqyc = kqyc;
	}
	public int getWeishua() {
		return weishua;
	}
	public void setWeishua(int weishua) {
		this.weishua = weishua;
	}
	public int getChidao() {
		return chidao;
	}
	public void setChidao(int chidao) {
		this.chidao = chidao;
	}
	public int getZaotui() {
		return zaotui;
	}
	public void setZaotui(int zaotui) {
		this.zaotui = zaotui;
	}
	public String getHege() {
		return hege;
	}
	public void setHege(String hege) {
		this.hege = hege;
	}
	public int getMqDays() {
		return mqDays;
	}
	public void setMqDays(int mqDays) {
		this.mqDays = mqDays;
	}
//	public int getMqWeeks() {
//		return mqWeeks;
//	}
//	public void setMqWeeks(int mqWeeks) {
//		this.mqWeeks = mqWeeks;
//	}
	public String getZhiban() {
		return zhiban;
	}
	public void setZhiban(String zhiban) {
		this.zhiban = zhiban;
	}
	
}
