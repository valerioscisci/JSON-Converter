package json_converter;

import java.awt.*;
import java.awt.event.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

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
			JSONObject jsonObj = (JSONObject) parser.parse(new FileReader(path));
			
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
		
		String[] headerTitles = {"Submited",
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
				"o) Fighting poverty",
				"social inclusion and social services",
				"emergency services",
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
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
				"Piloting of new solutions",
				"Adoption or transfer of technological solutions",
				"Material investment",
				"Other",
				"Common research",
				"Common strategies/action plans",
				"Joint learning",
				"New network",
				"clusters or innovation platforms",
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
		String rec = String.valueOf(record);
		rec = rec.replace("{", "");
		rec = rec.replace("}", "");
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
			}
		}
	}
}