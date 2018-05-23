package json_converter;

import java.awt.*;
import java.awt.event.*;
import javax.swing.*;
import javax.swing.GroupLayout.Alignment;
import javax.swing.LayoutStyle.ComponentPlacement;
import javax.swing.UIManager.*;

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
				
				JButton btnExit = new JButton("Exit");
				
				JFileChooser fileChooser = new JFileChooser();
				
				JLabel lblSelezionaIlFile = new JLabel("Select the file you want to convert:");
				lblSelezionaIlFile.setFont(new Font("Arial", Font.PLAIN, 14));
				
				JButton btnConvert = new JButton("Convert");
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
									.addPreferredGap(ComponentPlacement.RELATED, 412, Short.MAX_VALUE)
									.addComponent(btnExit, GroupLayout.PREFERRED_SIZE, 70, GroupLayout.PREFERRED_SIZE))
								.addGroup(groupLayout.createSequentialGroup()
									.addGap(199)
									.addComponent(lblJsonConverter, GroupLayout.PREFERRED_SIZE, 181, GroupLayout.PREFERRED_SIZE))
								.addGroup(groupLayout.createSequentialGroup()
									.addContainerGap()
									.addComponent(fileChooser, GroupLayout.DEFAULT_SIZE, 572, Short.MAX_VALUE)))
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
							.addPreferredGap(ComponentPlacement.RELATED, 143, Short.MAX_VALUE)
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