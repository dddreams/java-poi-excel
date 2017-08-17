package com.shure.entity;

/**
 * 图书类
 * @author shure
 *
 */
public class Book {
	private String bookId;
	private String cId;
	private String name;
	private String author;
	private String press;
	private int price;
	private String pressTime;
	
	public Book(){}

	public Book(String bookId, String name, String author, String press, int price, String pressTime, String cId) {
		super();
		this.bookId = bookId;
		this.name = name;
		this.author = author;
		this.press = press;
		this.price = price;
		this.pressTime = pressTime;
		this.cId = cId;
	}

	public String getBookId() {
		return bookId;
	}

	public void setBookId(String bookId) {
		this.bookId = bookId;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getAuthor() {
		return author;
	}

	public void setAuthor(String author) {
		this.author = author;
	}

	public String getPress() {
		return press;
	}

	public void setPress(String press) {
		this.press = press;
	}

	public int getPrice() {
		return price;
	}

	public void setPrice(int price) {
		this.price = price;
	}

	public String getPressTime() {
		return pressTime;
	}

	public void setPressTime(String pressTime) {
		this.pressTime = pressTime;
	}
	
	public String getCId(){
		return cId;
	}
	
	public void setCId(String cId){
		this.cId = cId;
	}
	
}
