package json_converter;

import java.awt.*;
import java.awt.event.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.sql.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Instant;
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
			for (int i=0; i<Integer.parseInt(survey_num); i++) {
				list.add(survey_record.get(i).toString());
			}
			
			/* Writes the xlsx file 
			 * - Start - */
			
			Calendar cal = Calendar.getInstance();
	        cal.setTime(Date.from(Instant.now()));
			String excelFileName = System.getProperty("user.home") + "/Desktop/" + String.format(
	                "Results_Cross_Border_Cooperation_%1$tY_%1$tm_%1$td_%1$tk_%1$tS_%1$tp.xlsx", cal); //name of the excel file
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
				if (jsonKey.equals("HappendAt")) {
					jsonValue = jsonValue.substring(7, 20);
					java.util.Date subTime = new java.util.Date(Long.parseLong(jsonValue));
					DateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss z");
					df.setTimeZone(java.util.TimeZone.getTimeZone("GMT+1")); 
					jsonValue = df.format(subTime);
				}
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
        String[] keys = {"HappendAt",
        		"name",
        		"surname",
        		"organisation",
        		"email",
        		"1_1",
        		"a2",
        		"b2",
        		"c2",
        		"d2",
        		"e2",
        		"f2",
        		"g2",
        		"h2",
        		"i2",
        		"l2",
        		"m2",
        		"n2",
        		"o2",
        		"p2",
        		"q2",
        		"r2",
        		"s2",
        		"t2",
        		"u2",
        		"v2",
        		"a31",
        		"a32",
        		"a33",
        		"a34",
        		"a35",
        		"a36",
        		"a37",
        		"b31",
        		"b32",
        		"b33",
        		"b34",
        		"b35",
        		"b36",
        		"b37",
        		"c31",
        		"c32",
        		"c33",
        		"c34",
        		"c35",
        		"c36",
        		"c37",
        		"d31",
        		"d32",
        		"d33",
        		"d34",
        		"d35",
        		"d36",
        		"d37",
        		"e31",
        		"e32",
        		"e33",
        		"e34",
        		"e35",
        		"e36",
        		"e37",
        		"f31",
        		"f32",
        		"f33",
        		"f34",
        		"f35",
        		"f36",
        		"f37",
        		"g31",
        		"g32",
        		"g33",
        		"g34",
        		"g35",
        		"g36",
        		"g37",
        		"h31",
        		"h32",
        		"h33",
        		"h34",
        		"h35",
        		"h36",
        		"h37",
        		"i31",
        		"i32",
        		"i33",
        		"i34",
        		"i35",
        		"i36",
        		"i37",
        		"l31",
        		"l32",
        		"l33",
        		"l34",
        		"l35",
        		"l36",
        		"l37",
        		"m31",
        		"m32",
        		"m33",
        		"m34",
        		"m35",
        		"m36",
        		"m37",
        		"n31",
        		"n32",
        		"n33",
        		"n34",
        		"n35",
        		"n36",
        		"n37",
        		"o31",
        		"o32",
        		"o33",
        		"o34",
        		"o35",
        		"o36",
        		"o37",
        		"p31",
        		"p32",
        		"p33",
        		"p34",
        		"p35",
        		"p36",
        		"p37",
        		"q31",
        		"q32",
        		"q33",
        		"q34",
        		"q35",
        		"q36",
        		"q37",
        		"r31",
        		"r32",
        		"r33",
        		"r34",
        		"r35",
        		"r36",
        		"r37",
        		"s31",
        		"s32",
        		"s33",
        		"s34",
        		"s35",
        		"s36",
        		"s37",
        		"t31",
        		"t32",
        		"t33",
        		"t34",
        		"t35",
        		"t36",
        		"t37",
        		"u31",
        		"u32",
        		"u33",
        		"u34",
        		"u35",
        		"u36",
        		"u37",
        		"v31",
        		"v32",
        		"v33",
        		"v34",
        		"v35",
        		"v36",
        		"v37",
        		"a4",
        		"b4",
        		"a51",
        		"a521",
        		"a522",
        		"a523",
        		"a524",
        		"a525",
        		"a526",
        		"b51",
        		"b521",
        		"b522",
        		"b523",
        		"b524",
        		"b525",
        		"b526",
        		"c51",
        		"c521",
        		"c522",
        		"c523",
        		"c524",
        		"c525",
        		"c526",
        		"d51",
        		"d521",
        		"d522",
        		"d523",
        		"d524",
        		"d525",
        		"d526",
        		"e51",
        		"e521",
        		"e522",
        		"e523",
        		"e524",
        		"e525",
        		"e526",
        		"f51",
        		"f521",
        		"f522",
        		"f523",
        		"f524",
        		"f525",
        		"f526",
        		"g51",
        		"g521",
        		"g522",
        		"g523",
        		"g524",
        		"g525",
        		"g526",
        		"h51",
        		"h521",
        		"h522",
        		"h523",
        		"h524",
        		"h525",
        		"h526",
        		"i51",
        		"i521",
        		"i522",
        		"i523",
        		"i524",
        		"i525",
        		"i526",
        		"l51",
        		"l521",
        		"l522",
        		"l523",
        		"l524",
        		"l525",
        		"l526",
        		"m51",
        		"m521",
        		"m522",
        		"m523",
        		"m524",
        		"m525",
        		"m526",
        		"n51",
        		"n521",
        		"n522",
        		"n523",
        		"n524",
        		"n525",
        		"n526",
        		"o51",
        		"o521",
        		"o522",
        		"o523",
        		"o524",
        		"o525",
        		"o526",
        		"p51",
        		"p521",
        		"p522",
        		"p523",
        		"p524",
        		"p525",
        		"p526",
        		"q51",
        		"q521",
        		"q522",
        		"q523",
        		"q524",
        		"q525",
        		"q526",
        		"r51",
        		"r521",
        		"r522",
        		"r523",
        		"r524",
        		"r525",
        		"r526",
        		"s51",
        		"s521",
        		"s522",
        		"s523",
        		"s524",
        		"s525",
        		"s526",
        		"t51",
        		"t521",
        		"t522",
        		"t523",
        		"t524",
        		"t525",
        		"t526",
        		"u51",
        		"u521",
        		"u522",
        		"u523",
        		"u524",
        		"u525",
        		"u526",
        		"v51",
        		"v521",
        		"v522",
        		"v523",
        		"v524",
        		"v525",
        		"v526",
        		"a61",
        		"a62",
        		"a63",
        		"a64",
        		"a65",
        		"a66",
        		"a67",
        		"a68",
        		"b61",
        		"b62",
        		"b63",
        		"b64",
        		"b65",
        		"b66",
        		"b67",
        		"b68",
        		"c61",
        		"c62",
        		"c63",
        		"c64",
        		"c65",
        		"c66",
        		"c67",
        		"c68",
        		"d61",
        		"d62",
        		"d63",
        		"d64",
        		"d65",
        		"d66",
        		"d67",
        		"d68",
        		"e61",
        		"e62",
        		"e63",
        		"e64",
        		"e65",
        		"e66",
        		"e67",
        		"e68",
        		"f61",
        		"f62",
        		"f63",
        		"f64",
        		"f65",
        		"f66",
        		"f67",
        		"f68",
        		"g61",
        		"g62",
        		"g63",
        		"g64",
        		"g65",
        		"g66",
        		"g67",
        		"g68",
        		"h61",
        		"h62",
        		"h63",
        		"h64",
        		"h65",
        		"h66",
        		"h67",
        		"h68",
        		"i61",
        		"i62",
        		"i63",
        		"i64",
        		"i65",
        		"i66",
        		"i67",
        		"i68",
        		"l61",
        		"l62",
        		"l63",
        		"l64",
        		"l65",
        		"l66",
        		"l67",
        		"l68",
        		"m61",
        		"m62",
        		"m63",
        		"m64",
        		"m65",
        		"m66",
        		"m67",
        		"m68",
        		"n61",
        		"n62",
        		"n63",
        		"n64",
        		"n65",
        		"n66",
        		"n67",
        		"n68",
        		"o61",
        		"o62",
        		"o63",
        		"o64",
        		"o65",
        		"o66",
        		"o67",
        		"o68",
        		"p61",
        		"p62",
        		"p63",
        		"p64",
        		"p65",
        		"p66",
        		"p67",
        		"p68",
        		"q61",
        		"q62",
        		"q63",
        		"q64",
        		"q65",
        		"q66",
        		"q67",
        		"q68",
        		"r61",
        		"r62",
        		"r63",
        		"r64",
        		"r65",
        		"r66",
        		"r67",
        		"r68",
        		"s61",
        		"s62",
        		"s63",
        		"s64",
        		"s65",
        		"s66",
        		"s67",
        		"s68",
        		"t61",
        		"t62",
        		"t63",
        		"t64",
        		"t65",
        		"t66",
        		"t67",
        		"t68",
        		"u61",
        		"u62",
        		"u63",
        		"u64",
        		"u65",
        		"u66",
        		"u67",
        		"u68",
        		"v61",
        		"v62",
        		"v63",
        		"v64",
        		"v65",
        		"v66",
        		"v67",
        		"v68",
        		"a7",
        		"a81",
        		"a82",
        		"a83",
        		"a84",
        		"a85",
        		"a86",
        		"a87",
        		"b81",
        		"b82",
        		"b83",
        		"b84",
        		"b85",
        		"b86",
        		"b87",
        		"c81",
        		"c82",
        		"c83",
        		"c84",
        		"c85",
        		"c86",
        		"c87",
        		"d81",
        		"d82",
        		"d83",
        		"d84",
        		"d85",
        		"d86",
        		"d87",
        		"e81",
        		"e82",
        		"e83",
        		"e84",
        		"e85",
        		"e86",
        		"e87",
        		"f81",
        		"f82",
        		"f83",
        		"f84",
        		"f85",
        		"f86",
        		"f87",
        		"g81",
        		"g82",
        		"g83",
        		"g84",
        		"g85",
        		"g86",
        		"g87",
        		"h81",
        		"h82",
        		"h83",
        		"h84",
        		"h85",
        		"h86",
        		"h87",
        		"i81",
        		"i82",
        		"i83",
        		"i84",
        		"i85",
        		"i86",
        		"i87",
        		"l81",
        		"l82",
        		"l83",
        		"l84",
        		"l85",
        		"l86",
        		"l87",
        		"m81",
        		"m82",
        		"m83",
        		"m84",
        		"m85",
        		"m86",
        		"m87",
        		"n81",
        		"n82",
        		"n83",
        		"n84",
        		"n85",
        		"n86",
        		"n87",
        		"o81",
        		"o82",
        		"o83",
        		"o84",
        		"o85",
        		"o86",
        		"o87",
        		"p81",
        		"p82",
        		"p83",
        		"p84",
        		"p85",
        		"p86",
        		"p87",
        		"q81",
        		"q82",
        		"q83",
        		"q84",
        		"q85",
        		"q86",
        		"q87",
        		"r81",
        		"r82",
        		"r83",
        		"r84",
        		"r85",
        		"r86",
        		"r87",
        		"s81",
        		"s82",
        		"s83",
        		"s84",
        		"s85",
        		"s86",
        		"s87",
        		"t81",
        		"t82",
        		"t83",
        		"t84",
        		"t85",
        		"t86",
        		"t87",
        		"u81",
        		"u82",
        		"u83",
        		"u84",
        		"u85",
        		"u86",
        		"u87",
        		"v81",
        		"v82",
        		"v83",
        		"v84",
        		"v85",
        		"v86",
        		"v87",
        		"a9",
        		"b9"};
        int i = 0;
        for (String key : keys) {
        	myMap.put(key, i);
        	i++;
        }
        return myMap;
    }
	
}