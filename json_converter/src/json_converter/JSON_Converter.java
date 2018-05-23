package json_converter;

import java.awt.*;
import java.awt.event.*;
import java.io.File;

import javax.swing.*;
import javax.swing.GroupLayout.Alignment;
import javax.swing.LayoutStyle.ComponentPlacement;
import javax.swing.UIManager.*;
import javax.swing.filechooser.FileFilter;

public class JSON_Converter{

	public static void main(String[] args) {
	    SwingUtilities.invokeLater(new Runnable() {
	    	
	        public void run() {
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
				
				JLabel lblNewLabel_1 = new JLabel("Selected file Name:");
				
				JLabel lblPath = new JLabel("");
				
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
					           return filename.endsWith(".json") || filename.endsWith(".json") ;
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
					groupLayout.createParallelGroup(Alignment.LEADING)
						.addGroup(groupLayout.createSequentialGroup()
							.addGroup(groupLayout.createParallelGroup(Alignment.LEADING)
								.addGroup(groupLayout.createSequentialGroup()
									.addContainerGap()
									.addComponent(lblSelezionaIlFile))
								.addGroup(groupLayout.createSequentialGroup()
									.addContainerGap()
									.addComponent(btnConvert)
									.addPreferredGap(ComponentPlacement.RELATED, 432, Short.MAX_VALUE)
									.addComponent(btnExit, GroupLayout.PREFERRED_SIZE, 70, GroupLayout.PREFERRED_SIZE))
								.addGroup(groupLayout.createSequentialGroup()
									.addGap(199)
									.addComponent(lblJsonConverter, GroupLayout.DEFAULT_SIZE, 181, Short.MAX_VALUE)
									.addGap(198))
								.addGroup(groupLayout.createSequentialGroup()
									.addContainerGap()
									.addGroup(groupLayout.createParallelGroup(Alignment.LEADING)
										.addComponent(fileChooser, GroupLayout.DEFAULT_SIZE, 572, Short.MAX_VALUE)
										.addGroup(groupLayout.createSequentialGroup()
											.addGap(6)
											.addGroup(groupLayout.createParallelGroup(Alignment.LEADING)
												.addGroup(groupLayout.createSequentialGroup()
													.addComponent(lblNewLabel_1)
													.addPreferredGap(ComponentPlacement.RELATED)
													.addComponent(lblName))
												.addGroup(groupLayout.createSequentialGroup()
													.addComponent(lblNewLabel)
													.addPreferredGap(ComponentPlacement.RELATED)
													.addComponent(lblPath)))))))
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
							.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
								.addComponent(lblNewLabel)
								.addComponent(lblPath))
							.addPreferredGap(ComponentPlacement.UNRELATED)
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
				frame.setSize(600,550);
				frame.setLocationRelativeTo(null);
	            frame.setVisible(true);
	            
	            btnExit.addActionListener(new ActionListener() {
	                public void actionPerformed(ActionEvent e)
	                {
	                   frame.dispose();
	                }
	            });
	        }
	    });
	    
	}
}