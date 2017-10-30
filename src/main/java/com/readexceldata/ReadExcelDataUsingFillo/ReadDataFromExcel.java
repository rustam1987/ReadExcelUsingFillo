package com.readexceldata.ReadExcelDataUsingFillo;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import com.codoid.products.exception.FilloException;
import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;

public class ReadDataFromExcel {

	public static void main(String[] args) throws FilloException {
		
		Fillo fillo=new Fillo();
		String dir=System.getProperty("user.dir");
		String query="Select * from userdata";
        Connection con=fillo.getConnection(System.getProperty("user.dir")+"\\Userdata.xlsx");
		Recordset rs=con.executeQuery(query);
		//Fetch data from excel sheet put Hash map
		Map<Integer,String> map=new HashMap<Integer,String>();
		
		
		while(rs.next()){
			System.out.println(rs.getField("First_Name"));
			String id=rs.getField("User_ID");
			int nid=Integer.parseInt(id);
			map.put(nid,rs.getField("First_Name")+" "+rs.getField("Last_Name") );
			
		}
        rs.close();
        con.close();
		for(Map.Entry<Integer,String> mp:map.entrySet()){
			System.out.println("User_Id :"+mp.getKey()+" User_FirstName and Last_Name :"+mp.getValue());
			
		}
		String excelNm="Userdata.xlsx";
		String query1="Select * from userdata";
		String flname="Last_Name";
		
		ArrayList <String>arr=readExcelData(excelNm, query1, flname);
		System.out.println(arr);
	}
	
	public static ArrayList<String> readExcelData(String ExcelName,String query,String feildName) throws FilloException{
		ArrayList<String> arr=new ArrayList<String>();
		Fillo fillo=new Fillo();
		String data="";
		Connection con=fillo.getConnection(System.getProperty("user.dir")+"\\"+ExcelName);
		Recordset rs=con.executeQuery(query);
		while(rs.next()){
		rs.getField(feildName);
		arr.add(rs.getField(feildName));
		
		}
		rs.close();
		con.close();
		return arr;
		
	}
	
	public String getCellData(String ExcelName,String query,String columnName) throws FilloException{
		Fillo fillo=new Fillo();
		Connection con=fillo.getConnection(ExcelName);
		Recordset rs=con.executeQuery(query);
		rs.next();
		return rs.getField(columnName);
		
	}
	
	
	
}
