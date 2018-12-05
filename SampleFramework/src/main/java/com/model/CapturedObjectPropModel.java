package com.model;
/**
 * The CapturedObjectPropModel class is used as a bean class to get the Captured
 * objects properties in the excel  sheet
 *
 * 
 */
public class CapturedObjectPropModel {
	String page=null;
	String name=null;
	String property=null;
	String value=null;
	
	public void setPage(String page){
		this.page=page;
	}
	
	public String getPage(){
		return page;
	}
	
	public void setName(String name){
		this.name=name;
	}
	
	public String getName(){
		return name;
	}
	
	public void setProperty(String property){
		this.property=property;
	}
	
	public String getProperty(){
		return property;
	}
	
	public void setValue(String value){
		this.value=value;
	}
	
	public String getValue(){
		return value;
	}
}
