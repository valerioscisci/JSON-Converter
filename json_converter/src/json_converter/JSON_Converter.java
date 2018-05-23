package json_converter;

import java.awt.*;
import javax.swing.*;

public class JSON_Converter{

	public static void main(String[] args) {
	    SwingUtilities.invokeLater(new Runnable() {
	    	
	        public void run() {
	        	JFrame frame = new JFrame("JSON Converter");
	        	JFrame.setDefaultLookAndFeelDecorated(true);
	            frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	            JLabel pane = new JLabel("Image and Text", JLabel.CENTER);
				frame.getContentPane().add(pane, BorderLayout.CENTER);
	            frame.pack();
	            frame.setVisible(true);
	        }
	    });
	}
	
}