package json_converter;

import java.awt.*;
import java.awt.event.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.sql.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Map;

import javax.swing.*;
import javax.swing.GroupLayout.Alignment;
import javax.swing.LayoutStyle.ComponentPlacement;
import javax.swing.UIManager.*;
import javax.swing.filechooser.FileFilter;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class JSON_Converter{
	
	public static void main(String[] args) {
	    SwingUtilities.invokeLater(new Runnable() {
	    	
	        public void run() {
	        	
	        	/* Creates the Interface */
	        	
	        	JFrame frame = new JFrame("JSON Converter");
	        	try {
	        	    for (LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
	        	        if ("Nimbus".equals(info.getName())) {
	        	            UIManager.setLookAndFeel(info.getClassName());
	        	            break;
	        	        }
	        	    }
	        	} catch (Exception e) {
	        		JFrame.setDefaultLookAndFeelDecorated(true);
	        	}
	            frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
				
				JLabel lblJsonConverter = new JLabel("JSON Converter");
				lblJsonConverter.setHorizontalAlignment(SwingConstants.CENTER);
				lblJsonConverter.setFont(new Font("SansSerif", Font.BOLD, 17));
				
				JButton btnConvert = new JButton("Convert");
				
				JLabel lblNewLabel = new JLabel("Selected file Path:");
				
				JLabel lblPath = new JLabel("");;
				
				JLabel lblNewLabel_1 = new JLabel("Selected file Name:");
				
				JLabel lblName = new JLabel("");
				
				JButton btnExit = new JButton("Exit");
				
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileFilter(new FileFilter() {

					   public String getDescription() {
					       return "JSON Files (*.json)";
					   }

					   public boolean accept(File f) {
					       if (f.isDirectory()) {
					           return true;
					       } else {
					           String filename = f.getName().toLowerCase();
					           return filename.endsWith(".json");
					       }
					   }
					});
				 fileChooser.addActionListener(new ActionListener() {
				      public void actionPerformed(ActionEvent e) {
				    	  lblPath.setText(fileChooser.getSelectedFile().getAbsolutePath());
				    	  lblName.setText(fileChooser.getSelectedFile().getName());
				      }
				    });
				
				JLabel lblSelezionaIlFile = new JLabel("Select the file you want to convert:");
				lblSelezionaIlFile.setFont(new Font("Arial", Font.PLAIN, 14));
				
				GroupLayout groupLayout = new GroupLayout(frame.getContentPane());
				groupLayout.setHorizontalGroup(
					groupLayout.createParallelGroup(Alignment.TRAILING)
						.addGroup(groupLayout.createSequentialGroup()
							.addGroup(groupLayout.createParallelGroup(Alignment.LEADING)
								.addGroup(Alignment.TRAILING, groupLayout.createSequentialGroup()
									.addContainerGap()
									.addComponent(fileChooser, GroupLayout.DEFAULT_SIZE, 622, Short.MAX_VALUE))
								.addGroup(groupLayout.createSequentialGroup()
									.addContainerGap()
									.addComponent(lblSelezionaIlFile))
								.addGroup(groupLayout.createSequentialGroup()
									.addGap(199)
									.addComponent(lblJsonConverter, GroupLayout.DEFAULT_SIZE, 231, Short.MAX_VALUE)
									.addGap(198))
								.addGroup(Alignment.TRAILING, groupLayout.createSequentialGroup()
									.addContainerGap()
									.addComponent(btnConvert)
									.addPreferredGap(ComponentPlacement.RELATED, 482, Short.MAX_VALUE)
									.addComponent(btnExit, GroupLayout.PREFERRED_SIZE, 70, GroupLayout.PREFERRED_SIZE))
								.addGroup(groupLayout.createSequentialGroup()
									.addContainerGap()
									.addComponent(lblNewLabel)
									.addPreferredGap(ComponentPlacement.RELATED)
									.addComponent(lblPath, GroupLayout.DEFAULT_SIZE, 524, Short.MAX_VALUE))
								.addGroup(groupLayout.createSequentialGroup()
									.addContainerGap()
									.addComponent(lblNewLabel_1)
									.addPreferredGap(ComponentPlacement.RELATED)
									.addComponent(lblName, GroupLayout.DEFAULT_SIZE, 515, Short.MAX_VALUE)))
							.addContainerGap())
				);
				groupLayout.setVerticalGroup(
					groupLayout.createParallelGroup(Alignment.LEADING)
						.addGroup(groupLayout.createSequentialGroup()
							.addComponent(lblJsonConverter)
							.addGap(27)
							.addComponent(lblSelezionaIlFile)
							.addPreferredGap(ComponentPlacement.RELATED)
							.addComponent(fileChooser, GroupLayout.PREFERRED_SIZE, 261, GroupLayout.PREFERRED_SIZE)
							.addPreferredGap(ComponentPlacement.RELATED)
							.addGroup(groupLayout.createParallelGroup(Alignment.LEADING)
								.addComponent(lblPath, GroupLayout.PREFERRED_SIZE, 16, GroupLayout.PREFERRED_SIZE)
								.addComponent(lblNewLabel))
							.addGap(12)
							.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
								.addComponent(lblNewLabel_1)
								.addComponent(lblName))
							.addPreferredGap(ComponentPlacement.RELATED, 93, Short.MAX_VALUE)
							.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
								.addComponent(btnExit)
								.addComponent(btnConvert))
							.addContainerGap())
				);
				frame.getContentPane().setLayout(groupLayout);
				frame.setSize(650,550);
				frame.setLocationRelativeTo(null);
	            frame.setVisible(true);
	            
	            btnConvert.addActionListener(new ActionListener() {
	                public void actionPerformed(ActionEvent e)
	                {
	                   jsonToArray(lblPath.getText());
	                }
	            });
	            
	            btnExit.addActionListener(new ActionListener() {
	                public void actionPerformed(ActionEvent e)
	                {
	                   frame.dispose();
	                }
	            });
	        }
	    });
	    
	}
	
	/* Method used to parse the json file and then convert it to xlsx format */
	
	public static void jsonToArray(String path){
		JSONParser parser = new JSONParser();
		try {
			JSONObject jsonObj = (JSONObject) parser.parse(new InputStreamReader(new FileInputStream(path), "UTF-8"));
			
			String survey_num = String.valueOf(jsonObj.get("ResultCount"));
			
			JSONArray survey_record = (JSONArray) jsonObj.get("Data");
			ArrayList<String> list = new ArrayList<String>();
			for (int i=0; i<=survey_num.length(); i++) {
				list.add(survey_record.get(i).toString());
			} 
			
			/* Writes the xlsx file 
			 * - Start - */
			
			String excelFileName = System.getProperty("user.home") + "/Desktop/Results.xlsx"; //name of the excel file
			String sheetName = "Results"; //name of the sheet
			
			XSSFWorkbook wb = new XSSFWorkbook();
			XSSFSheet sheet = wb.createSheet(sheetName) ;
			
			/* Create the Header of the xlsx file */
			
			createHeader(sheet);
			
			
			
			// creating a row for each element in the list
			
			int i=1;
			
			for (Object rec : list)
			  {
				XSSFRow row = sheet.createRow(i);
				
				// creating a cell for each element
					
				sortElements(sheet, row, i, rec);
				
				i++;
			  }	 
			
			FileOutputStream fileOut = new FileOutputStream(excelFileName);

			//write this workbook to an Outputstream
			
			wb.write(fileOut);
			wb.close();
			fileOut.flush();
			fileOut.close();
			
			/* - End - */
			
    	} catch (FileNotFoundException e) {
    		e.printStackTrace();
    	} catch (IOException e) {
    		e.printStackTrace();
    	} catch (ParseException e) {
    		e.printStackTrace();
    	}
	}
	
	/* Method to create the Header of the xlsx file */
	
	public static void createHeader(XSSFSheet xls) {
		XSSFRow row = xls.createRow(0);
		
		String[] headerTitles = {"Submitted",
				"Name",
				"Surname",
				"Organisation",
				"E-mail Address",
				"1.1) What type of organisation do you work for?",
				"a) Scientific cooperation",
				"b) Research and innovation",
				"c) ICT infrastructure and services",
				"d) SMEs development and internationalisation",
				"e) Financial services and credit market",
				"f) Low-carbon economy and renewable energy",
				"g) Climate change and risk management",
				"h) Protection and valorisation of cultural heritage",
				"i) Environmental protection and preservation",
				"l) Resource-efficiency and blue growth",
				"m) Transport connections",
				"n) Labour market and labour mobility",
				"o) Fighting poverty social inclusion and social services emergency services",
				"p) Education and students' mobility",
				"q) Administrative capacity",
				"r) Promotion of human rights and fundamental freedoms",
				"s) Market integration",
				"t) Migration and mobility of people",
				"u) Confidence building and reducing cultural barriers",
				"v) Safety and security",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"Common experience",
				"Common resource",
				"Common infrastructure",
				"Common challenge",
				"Common research expertise",
				"Geographical proximity",
				"Other",
				"a) your organization",
				"b) your local area",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"To what extent?",
				"Economic growth",
				"Jobs created",
				"Sustainable management of resources",
				"Decrease in poverty and social exclusion",
				"Improved capacity in public bodies",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"a) describe aspects",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"National Institutions or Agencies",
				"Regional Government",
				"Local Government",
				"Business",
				"Education",
				"Third Sector",
				"Other",
				"a) are currently in place?",
				"b) should be in place in the future?"
};
		
		for(int i=0; i < headerTitles.length ; i++) {
			XSSFCell cell = row.createCell(i);
			cell.setCellValue(headerTitles[i]);
			xls.autoSizeColumn(i);
		}
	}
	
	/* Method that separates the json key/value pair and writes the value in the right column within the xlsx file */
	
	public static void sortElements(XSSFSheet xls, XSSFRow row, int row_num, Object record) {
		String jsonKey, jsonValue;
		String[] parts;
		String[] questions = {"2_1","a3","b3","c3","d3","e3","f3","g3","h3","i3","l3","m3","n3","o3","p3","q3","r3","s3","t3","u3","v3"};
		String[] question5 = {"a52","b52","c52","d52","e52","f52","g52","h52","i52","l52","m52","n52","o52","p52","q52","r52","s52","t52","u52","v52"};
		String[] question6 = {"a6","b6","c6","d6","e6","f6","g6","h6","i6","l6","m6","n6","o6","p6","q6","r6","s6","t6","u6","v6"};
		String[] question8 = {"a8","b8","c8","d8","e8","f8","g8","h8","i8","l8","m8","n8","o8","p8","q8","r8","s8","t8","u8","v8"};
 		int col_num = 0;
		String rec = String.valueOf(record);
		rec = rec.replace("{", "");
		rec = rec.replace("}", "");
		rec = rec.replace("\\\"", "");
		rec = rec.replace("\\n", " ");
		Map<String, Integer> mapping = createMap();
		while(!rec.equals("")) {
			parts = rec.split(":",2);
			rec = parts[1];
			jsonKey = parts[0];
			jsonKey = jsonKey.replace("\"", "");
			if(rec.substring(0, 1).equals("[")) {
				rec = rec.substring(1);
				parts = rec.split("]", 2);
				rec = parts[1];
				try {
					rec = rec.substring(1);
				} catch (Exception e) {
					rec = "";
				}
				jsonValue = parts[0];
				if (jsonKey.equals("HappendAt")) {
					jsonValue = jsonValue.substring(7, 21);
					java.util.Date subTime = new java.util.Date(Long.parseLong(jsonValue) * 1000);
					DateFormat df = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
					jsonValue = df.format(subTime);
				}
			} else {
				rec = rec.substring(1);
				parts = rec.split("\"", 2);
				rec = parts[1];
				try {
					rec = rec.substring(1);
				} catch (Exception e) {
					rec = "";
				}
				jsonValue = parts[0];
				jsonValue = jsonValue.replace("\"", "");
			}
			if (Arrays.asList(questions).contains(jsonKey) || Arrays.asList(question5).contains(jsonKey) || Arrays.asList(question6).contains(jsonKey) || Arrays.asList(question8).contains(jsonKey)) {
				jsonValue = jsonValue.replace("\"", "");
				parts = jsonValue.split(",");
				for (String val : parts) {
					try {
						col_num = mapping.get(val);
						createCell(xls, row, col_num, "X");
						} catch (Exception e) {
							if (Arrays.asList(questions).contains(jsonKey) || Arrays.asList(question8).contains(jsonKey)) {
								jsonKey = jsonKey + "7";
								col_num = mapping.get(jsonKey);
							} else if (Arrays.asList(question5).contains(jsonKey)){
								jsonKey = jsonKey + "6";
								col_num = mapping.get(jsonKey);
							} else if (Arrays.asList(question6).contains(jsonKey)){
								jsonKey = jsonKey + "8";
								col_num = mapping.get(jsonKey);
							}
							createCell(xls, row, col_num, val);
						}
				}
			} else {
				try {
				col_num = mapping.get(jsonKey);
				} catch (Exception e) {
					col_num= 0;
				}
				createCell(xls, row, col_num, jsonValue);
			}
		}
	}
	
	/* Creates a cell */
	
	private static void createCell(XSSFSheet x, XSSFRow r, int c, String v){
		XSSFCell cell = r.createCell(c);
		cell.setCellValue(v);
		x.autoSizeColumn(c);
	}
	
	/* Creates the mapping between json keys and xlsx column indexes */
	
	private static Map<String, Integer> createMap()
    {
        Map<String,Integer> myMap = new HashMap<String,Integer>();
        myMap.put("HappendAt", 0);
        myMap.put("name", 1);
        myMap.put("surname", 2);
        myMap.put("organisation", 3);
        myMap.put("email", 4);
        myMap.put("1_1", 5);
        myMap.put("a2", 6);
        myMap.put("b2", 7);
        myMap.put("c2", 8);
        myMap.put("d2", 9);
        myMap.put("e2", 10);
        myMap.put("f2", 11);
        myMap.put("g2", 12);
        myMap.put("h2", 13);
        myMap.put("i2", 14);
        myMap.put("l2", 15);
        myMap.put("m2", 16);
        myMap.put("n2", 17);
        myMap.put("o2", 18);
        myMap.put("p2", 19);
        myMap.put("q2", 20);
        myMap.put("r2", 21);
        myMap.put("s2", 22);
        myMap.put("t2", 23);
        myMap.put("u2", 24);
        myMap.put("v2", 25);
        myMap.put("a31", 26);
        myMap.put("a32", 27);
        myMap.put("a33", 28);
        myMap.put("a34", 29);
        myMap.put("a35", 30);
        myMap.put("a36", 31);
        myMap.put("a37", 32);
        myMap.put("b31", 33);
        myMap.put("b32", 34);
        myMap.put("b33", 35);
        myMap.put("b34", 36);
        myMap.put("b35", 37);
        myMap.put("b36", 38);
        myMap.put("b37", 39);
        myMap.put("c31", 40);
        myMap.put("c32", 41);
        myMap.put("c33", 42);
        myMap.put("c34", 43);
        myMap.put("c35", 44);
        myMap.put("c36", 45);
        myMap.put("c37", 46);
        myMap.put("d31", 47);
        myMap.put("d32", 48);
        myMap.put("d33", 49);
        myMap.put("d34", 50);
        myMap.put("d35", 51);
        myMap.put("d36", 52);
        myMap.put("d37", 53);
        myMap.put("e31", 54);
        myMap.put("e32", 55);
        myMap.put("e33", 56);
        myMap.put("e34", 57);
        myMap.put("e35", 58);
        myMap.put("e36", 59);
        myMap.put("e37", 60);
        myMap.put("f31", 61);
        myMap.put("f32", 62);
        myMap.put("f33", 63);
        myMap.put("f34", 64);
        myMap.put("f35", 65);
        myMap.put("f36", 66);
        myMap.put("f37", 67);
        myMap.put("g31", 68);
        myMap.put("g32", 69);
        myMap.put("g33", 70);
        myMap.put("g34", 71);
        myMap.put("g35", 72);
        myMap.put("g36", 73);
        myMap.put("g37", 74);
        myMap.put("h31", 75);
        myMap.put("h32", 76);
        myMap.put("h33", 77);
        myMap.put("h34", 78);
        myMap.put("h35", 79);
        myMap.put("h36", 80);
        myMap.put("h37", 81);
        myMap.put("i31", 82);
        myMap.put("i32", 83);
        myMap.put("i33", 84);
        myMap.put("i34", 85);
        myMap.put("i35", 86);
        myMap.put("i36", 87);
        myMap.put("i37", 88);
        myMap.put("l31", 89);
        myMap.put("l32", 90);
        myMap.put("l33", 91);
        myMap.put("l34", 92);
        myMap.put("l35", 93);
        myMap.put("l36", 94);
        myMap.put("l37", 95);
        myMap.put("m31", 96);
        myMap.put("m32", 97);
        myMap.put("m33", 98);
        myMap.put("m34", 99);
        myMap.put("m35", 100);
        myMap.put("m36", 101);
        myMap.put("m37", 102);
        myMap.put("n31", 103);
        myMap.put("n32", 104);
        myMap.put("n33", 105);
        myMap.put("n34", 106);
        myMap.put("n35", 107);
        myMap.put("n36", 108);
        myMap.put("n37", 109);
        myMap.put("o31", 110);
        myMap.put("o32", 111);
        myMap.put("o33", 112);
        myMap.put("o34", 113);
        myMap.put("o35", 114);
        myMap.put("o36", 115);
        myMap.put("o37", 116);
        myMap.put("p31", 117);
        myMap.put("p32", 118);
        myMap.put("p33", 119);
        myMap.put("p34", 120);
        myMap.put("p35", 121);
        myMap.put("p36", 122);
        myMap.put("p37", 123);
        myMap.put("q31", 124);
        myMap.put("q32", 125);
        myMap.put("q33", 126);
        myMap.put("q34", 127);
        myMap.put("q35", 128);
        myMap.put("q36", 129);
        myMap.put("q37", 130);
        myMap.put("r31", 131);
        myMap.put("r32", 132);
        myMap.put("r33", 133);
        myMap.put("r34", 134);
        myMap.put("r35", 135);
        myMap.put("r36", 136);
        myMap.put("r37", 137);
        myMap.put("s31", 138);
        myMap.put("s32", 139);
        myMap.put("s33", 140);
        myMap.put("s34", 141);
        myMap.put("s35", 142);
        myMap.put("s36", 143);
        myMap.put("s37", 144);
        myMap.put("t31", 145);
        myMap.put("t32", 146);
        myMap.put("t33", 147);
        myMap.put("t34", 148);
        myMap.put("t35", 149);
        myMap.put("t36", 150);
        myMap.put("t37", 151);
        myMap.put("u31", 152);
        myMap.put("u32", 153);
        myMap.put("u33", 154);
        myMap.put("u34", 155);
        myMap.put("u35", 156);
        myMap.put("u36", 157);
        myMap.put("u37", 158);
        myMap.put("v31", 159);
        myMap.put("v32", 160);
        myMap.put("v33", 161);
        myMap.put("v34", 162);
        myMap.put("v35", 163);
        myMap.put("v36", 164);
        myMap.put("v37", 165);
        myMap.put("a4", 166);
        myMap.put("b4", 167);
        myMap.put("a51", 168);
        myMap.put("a521", 169);
        myMap.put("a522", 170);
        myMap.put("a523", 171);
        myMap.put("a524", 172);
        myMap.put("a525", 173);
        myMap.put("a526", 174);
        myMap.put("b51", 175);
        myMap.put("b521", 176);
        myMap.put("b522", 177);
        myMap.put("b523", 178);
        myMap.put("b524", 179);
        myMap.put("b525", 180);
        myMap.put("b526", 181);
        myMap.put("c51", 182);
        myMap.put("c521", 183);
        myMap.put("c522", 184);
        myMap.put("c523", 185);
        myMap.put("c524", 186);
        myMap.put("c525", 187);
        myMap.put("c526", 188);
        myMap.put("d51", 189);
        myMap.put("d521", 190);
        myMap.put("d522", 191);
        myMap.put("d523", 192);
        myMap.put("d524", 193);
        myMap.put("d525", 194);
        myMap.put("d526", 195);
        myMap.put("e51", 196);
        myMap.put("e521", 197);
        myMap.put("e522", 198);
        myMap.put("e523", 199);
        myMap.put("e524", 200);
        myMap.put("e525", 201);
        myMap.put("e526", 202);
        myMap.put("f51", 203);
        myMap.put("f521", 204);
        myMap.put("f522", 205);
        myMap.put("f523", 206);
        myMap.put("f524", 207);
        myMap.put("f525", 208);
        myMap.put("f526", 209);
        myMap.put("g51", 210);
        myMap.put("g521", 211);
        myMap.put("g522", 212);
        myMap.put("g523", 213);
        myMap.put("g524", 214);
        myMap.put("g525", 215);
        myMap.put("g526", 216);
        myMap.put("h51", 217);
        myMap.put("h521", 218);
        myMap.put("h522", 219);
        myMap.put("h523", 220);
        myMap.put("h524", 221);
        myMap.put("h525", 222);
        myMap.put("h526", 223);
        myMap.put("i51", 224);
        myMap.put("i521", 225);
        myMap.put("i522", 226);
        myMap.put("i523", 227);
        myMap.put("i524", 228);
        myMap.put("i525", 229);
        myMap.put("i526", 230);
        myMap.put("l51", 231);
        myMap.put("l521", 232);
        myMap.put("l522", 233);
        myMap.put("l523", 234);
        myMap.put("l524", 235);
        myMap.put("l525", 236);
        myMap.put("l526", 237);
        myMap.put("m51", 238);
        myMap.put("m521", 239);
        myMap.put("m522", 240);
        myMap.put("m523", 241);
        myMap.put("m524", 242);
        myMap.put("m525", 243);
        myMap.put("m526", 244);
        myMap.put("n51", 245);
        myMap.put("n521", 246);
        myMap.put("n522", 247);
        myMap.put("n523", 248);
        myMap.put("n524", 249);
        myMap.put("n525", 250);
        myMap.put("n526", 251);
        myMap.put("o51", 252);
        myMap.put("o521", 253);
        myMap.put("o522", 254);
        myMap.put("o523", 255);
        myMap.put("o524", 256);
        myMap.put("o525", 257);
        myMap.put("o526", 258);
        myMap.put("p51", 259);
        myMap.put("p521", 260);
        myMap.put("p522", 261);
        myMap.put("p523", 262);
        myMap.put("p524", 263);
        myMap.put("p525", 264);
        myMap.put("p526", 265);
        myMap.put("q51", 266);
        myMap.put("q521", 267);
        myMap.put("q522", 268);
        myMap.put("q523", 269);
        myMap.put("q524", 270);
        myMap.put("q525", 271);
        myMap.put("q526", 272);
        myMap.put("r51", 273);
        myMap.put("r521", 274);
        myMap.put("r522", 275);
        myMap.put("r523", 276);
        myMap.put("r524", 277);
        myMap.put("r525", 278);
        myMap.put("r526", 279);
        myMap.put("s51", 280);
        myMap.put("s521", 281);
        myMap.put("s522", 282);
        myMap.put("s523", 283);
        myMap.put("s524", 284);
        myMap.put("s525", 285);
        myMap.put("s526", 286);
        myMap.put("t51", 287);
        myMap.put("t521", 288);
        myMap.put("t522", 289);
        myMap.put("t523", 290);
        myMap.put("t524", 291);
        myMap.put("t525", 292);
        myMap.put("t526", 293);
        myMap.put("u51", 294);
        myMap.put("u521", 295);
        myMap.put("u522", 296);
        myMap.put("u523", 297);
        myMap.put("u524", 298);
        myMap.put("u525", 299);
        myMap.put("u526", 300);
        myMap.put("v51", 301);
        myMap.put("v521", 302);
        myMap.put("v522", 303);
        myMap.put("v523", 304);
        myMap.put("v524", 305);
        myMap.put("v525", 306);
        myMap.put("v526", 307);
        myMap.put("a61", 308);
        myMap.put("a62", 309);
        myMap.put("a63", 310);
        myMap.put("a64", 311);
        myMap.put("a65", 312);
        myMap.put("a66", 313);
        myMap.put("a67", 314);
        myMap.put("a68", 315);
        myMap.put("b61", 316);
        myMap.put("b62", 317);
        myMap.put("b63", 318);
        myMap.put("b64", 319);
        myMap.put("b65", 320);
        myMap.put("b66", 321);
        myMap.put("b67", 322);
        myMap.put("b68", 323);
        myMap.put("c61", 324);
        myMap.put("c62", 325);
        myMap.put("c63", 326);
        myMap.put("c64", 327);
        myMap.put("c65", 328);
        myMap.put("c66", 329);
        myMap.put("c67", 330);
        myMap.put("c68", 331);
        myMap.put("d61", 332);
        myMap.put("d62", 333);
        myMap.put("d63", 334);
        myMap.put("d64", 335);
        myMap.put("d65", 336);
        myMap.put("d66", 337);
        myMap.put("d67", 338);
        myMap.put("d68", 339);
        myMap.put("e61", 340);
        myMap.put("e62", 341);
        myMap.put("e63", 342);
        myMap.put("e64", 343);
        myMap.put("e65", 344);
        myMap.put("e66", 345);
        myMap.put("e67", 346);
        myMap.put("e68", 347);
        myMap.put("f61", 348);
        myMap.put("f62", 349);
        myMap.put("f63", 350);
        myMap.put("f64", 351);
        myMap.put("f65", 352);
        myMap.put("f66", 353);
        myMap.put("f67", 354);
        myMap.put("f68", 355);
        myMap.put("g61", 356);
        myMap.put("g62", 357);
        myMap.put("g63", 358);
        myMap.put("g64", 359);
        myMap.put("g65", 360);
        myMap.put("g66", 361);
        myMap.put("g67", 362);
        myMap.put("g68", 363);
        myMap.put("h61", 364);
        myMap.put("h62", 365);
        myMap.put("h63", 366);
        myMap.put("h64", 367);
        myMap.put("h65", 368);
        myMap.put("h66", 369);
        myMap.put("h67", 370);
        myMap.put("h68", 371);
        myMap.put("i61", 372);
        myMap.put("i62", 373);
        myMap.put("i63", 374);
        myMap.put("i64", 375);
        myMap.put("i65", 376);
        myMap.put("i66", 377);
        myMap.put("i67", 378);
        myMap.put("i68", 379);
        myMap.put("l61", 380);
        myMap.put("l62", 381);
        myMap.put("l63", 382);
        myMap.put("l64", 383);
        myMap.put("l65", 384);
        myMap.put("l66", 385);
        myMap.put("l67", 386);
        myMap.put("l68", 387);
        myMap.put("m61", 388);
        myMap.put("m62", 389);
        myMap.put("m63", 390);
        myMap.put("m64", 391);
        myMap.put("m65", 392);
        myMap.put("m66", 393);
        myMap.put("m67", 394);
        myMap.put("m68", 395);
        myMap.put("n61", 396);
        myMap.put("n62", 397);
        myMap.put("n63", 398);
        myMap.put("n64", 399);
        myMap.put("n65", 400);
        myMap.put("n66", 401);
        myMap.put("n67", 402);
        myMap.put("n68", 403);
        myMap.put("o61", 404);
        myMap.put("o62", 405);
        myMap.put("o63", 406);
        myMap.put("o64", 407);
        myMap.put("o65", 408);
        myMap.put("o66", 409);
        myMap.put("o67", 410);
        myMap.put("o68", 411);
        myMap.put("p61", 412);
        myMap.put("p62", 413);
        myMap.put("p63", 414);
        myMap.put("p64", 415);
        myMap.put("p65", 416);
        myMap.put("p66", 417);
        myMap.put("p67", 418);
        myMap.put("p68", 419);
        myMap.put("q61", 420);
        myMap.put("q62", 421);
        myMap.put("q63", 422);
        myMap.put("q64", 423);
        myMap.put("q65", 424);
        myMap.put("q66", 425);
        myMap.put("q67", 426);
        myMap.put("q68", 427);
        myMap.put("r61", 428);
        myMap.put("r62", 429);
        myMap.put("r63", 430);
        myMap.put("r64", 431);
        myMap.put("r65", 432);
        myMap.put("r66", 433);
        myMap.put("r67", 434);
        myMap.put("r68", 435);
        myMap.put("s61", 436);
        myMap.put("s62", 437);
        myMap.put("s63", 438);
        myMap.put("s64", 439);
        myMap.put("s65", 440);
        myMap.put("s66", 441);
        myMap.put("s67", 442);
        myMap.put("s68", 443);
        myMap.put("t61", 444);
        myMap.put("t62", 445);
        myMap.put("t63", 446);
        myMap.put("t64", 447);
        myMap.put("t65", 448);
        myMap.put("t66", 449);
        myMap.put("t67", 450);
        myMap.put("t68", 451);
        myMap.put("u61", 452);
        myMap.put("u62", 453);
        myMap.put("u63", 454);
        myMap.put("u64", 455);
        myMap.put("u65", 456);
        myMap.put("u66", 457);
        myMap.put("u67", 458);
        myMap.put("u68", 459);
        myMap.put("v61", 460);
        myMap.put("v62", 461);
        myMap.put("v63", 462);
        myMap.put("v64", 463);
        myMap.put("v65", 464);
        myMap.put("v66", 465);
        myMap.put("v67", 466);
        myMap.put("v68", 467);
        myMap.put("a7", 468);
        myMap.put("a81", 469);
        myMap.put("a82", 470);
        myMap.put("a83", 471);
        myMap.put("a84", 472);
        myMap.put("a85", 473);
        myMap.put("a86", 474);
        myMap.put("a87", 475);
        myMap.put("b81", 476);
        myMap.put("b82", 477);
        myMap.put("b83", 478);
        myMap.put("b84", 479);
        myMap.put("b85", 480);
        myMap.put("b86", 481);
        myMap.put("b87", 482);
        myMap.put("c81", 483);
        myMap.put("c82", 484);
        myMap.put("c83", 485);
        myMap.put("c84", 486);
        myMap.put("c85", 487);
        myMap.put("c86", 488);
        myMap.put("c87", 489);
        myMap.put("d81", 490);
        myMap.put("d82", 491);
        myMap.put("d83", 492);
        myMap.put("d84", 493);
        myMap.put("d85", 494);
        myMap.put("d86", 495);
        myMap.put("d87", 496);
        myMap.put("e81", 497);
        myMap.put("e82", 498);
        myMap.put("e83", 499);
        myMap.put("e84", 500);
        myMap.put("e85", 501);
        myMap.put("e86", 502);
        myMap.put("e87", 503);
        myMap.put("f81", 504);
        myMap.put("f82", 505);
        myMap.put("f83", 506);
        myMap.put("f84", 507);
        myMap.put("f85", 508);
        myMap.put("f86", 509);
        myMap.put("f87", 510);
        myMap.put("g81", 511);
        myMap.put("g82", 512);
        myMap.put("g83", 513);
        myMap.put("g84", 514);
        myMap.put("g85", 515);
        myMap.put("g86", 516);
        myMap.put("g87", 517);
        myMap.put("h81", 518);
        myMap.put("h82", 519);
        myMap.put("h83", 520);
        myMap.put("h84", 521);
        myMap.put("h85", 522);
        myMap.put("h86", 523);
        myMap.put("h87", 524);
        myMap.put("i81", 525);
        myMap.put("i82", 526);
        myMap.put("i83", 527);
        myMap.put("i84", 528);
        myMap.put("i85", 529);
        myMap.put("i86", 530);
        myMap.put("i87", 531);
        myMap.put("l81", 532);
        myMap.put("l82", 533);
        myMap.put("l83", 534);
        myMap.put("l84", 535);
        myMap.put("l85", 536);
        myMap.put("l86", 537);
        myMap.put("l87", 538);
        myMap.put("m81", 539);
        myMap.put("m82", 540);
        myMap.put("m83", 541);
        myMap.put("m84", 542);
        myMap.put("m85", 543);
        myMap.put("m86", 544);
        myMap.put("m87", 545);
        myMap.put("n81", 546);
        myMap.put("n82", 547);
        myMap.put("n83", 548);
        myMap.put("n84", 549);
        myMap.put("n85", 550);
        myMap.put("n86", 551);
        myMap.put("n87", 552);
        myMap.put("o81", 553);
        myMap.put("o82", 554);
        myMap.put("o83", 555);
        myMap.put("o84", 556);
        myMap.put("o85", 557);
        myMap.put("o86", 558);
        myMap.put("o87", 559);
        myMap.put("p81", 560);
        myMap.put("p82", 561);
        myMap.put("p83", 562);
        myMap.put("p84", 563);
        myMap.put("p85", 564);
        myMap.put("p86", 565);
        myMap.put("p87", 566);
        myMap.put("q81", 567);
        myMap.put("q82", 568);
        myMap.put("q83", 569);
        myMap.put("q84", 570);
        myMap.put("q85", 571);
        myMap.put("q86", 572);
        myMap.put("q87", 573);
        myMap.put("r81", 574);
        myMap.put("r82", 575);
        myMap.put("r83", 576);
        myMap.put("r84", 577);
        myMap.put("r85", 578);
        myMap.put("r86", 579);
        myMap.put("r87", 580);
        myMap.put("s81", 581);
        myMap.put("s82", 582);
        myMap.put("s83", 583);
        myMap.put("s84", 584);
        myMap.put("s85", 585);
        myMap.put("s86", 586);
        myMap.put("s87", 587);
        myMap.put("t81", 588);
        myMap.put("t82", 589);
        myMap.put("t83", 590);
        myMap.put("t84", 591);
        myMap.put("t85", 592);
        myMap.put("t86", 593);
        myMap.put("t87", 594);
        myMap.put("u81", 595);
        myMap.put("u82", 596);
        myMap.put("u83", 597);
        myMap.put("u84", 598);
        myMap.put("u85", 599);
        myMap.put("u86", 600);
        myMap.put("u87", 601);
        myMap.put("v81", 602);
        myMap.put("v82", 603);
        myMap.put("v83", 604);
        myMap.put("v84", 605);
        myMap.put("v85", 606);
        myMap.put("v86", 607);
        myMap.put("v87", 608);
        myMap.put("a9", 609);
        myMap.put("b9", 610);
        return myMap;
    }
	
}