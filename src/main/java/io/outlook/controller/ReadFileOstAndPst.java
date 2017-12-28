package io.outlook.controller;

import com.pff.*;

import java.io.File;
import java.util.*;

import javax.swing.JFileChooser;

public class ReadFileOstAndPst {
	int depth = -1;

	public void App(String filename) {
		try {
			PSTFile pstFile = new PSTFile(filename);
			System.out.println(pstFile.getMessageStore().getDisplayName());
			processFolder(pstFile.getRootFolder());
		} catch (Exception err) {
			err.printStackTrace();
		}
	}

	public void processFolder(PSTFolder folder) throws PSTException, java.io.IOException {
		depth++;
		// the root folder doesn't have a display name
		if (depth > 0) {
			printDepth();
			System.out.println(folder.getDisplayName());
		}

		// go through the folders...
		if (folder.hasSubfolders()) {
			Vector<PSTFolder> childFolders = folder.getSubFolders();
			for (PSTFolder childFolder : childFolders) {
				processFolder(childFolder);
			}
		}

		// and now the emails for this folder
		if (folder.getContentCount() > 0) {
			depth++;
			PSTMessage email = (PSTMessage) folder.getNextChild();
			while (email != null) {
				printDepth();

				System.out.println("Email Subject: " + email.getSubject());
				System.out.println("Email Body: " + email.getBody());
				System.out.println("Email Body: " + email.getDisplayName());

				email = (PSTMessage) folder.getNextChild();
			}
			depth--;
		}
		depth--;
	}

	public void printDepth() {
		for (int x = 0; x < depth - 1; x++) {
			System.out.print(" | ");
		}
		System.out.print(" |- ");
	}

	public static void main(String[] args) {
		ReadFileOstAndPst readFileOstAndPst = new ReadFileOstAndPst();
		JFileChooser choose = new JFileChooser();
        File Folder = new File("D:\\eclipse-workspace\\outlook");
        choose.setCurrentDirectory(Folder);
        int showOpenDialog = choose.showOpenDialog(null);
        choose.setCurrentDirectory(new File(System.getProperty("user.home")));
        if (showOpenDialog == JFileChooser.APPROVE_OPTION) {
            File file = choose.getSelectedFile();
            String fileName =file.getName();
            readFileOstAndPst.App(fileName);
        }
		
		// final String finalName =
		// "C:/Venkat_06012014/Blog/eMails/pst/nallave_ost_format.ost";
//		JFileChooser chooser = new JFileChooser();
//		
//		chooser.showOpenDialog(chooser);
//		File file= chooser.getSelectedFile();
//		String fileName = file.getName();
		
		
	}

}
