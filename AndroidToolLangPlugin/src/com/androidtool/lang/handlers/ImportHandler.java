package com.androidtool.lang.handlers;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;

import org.eclipse.core.commands.ExecutionException;
import org.eclipse.core.resources.IProject;
import org.eclipse.core.resources.IResource;
import org.eclipse.core.runtime.CoreException;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.ui.IWorkbenchWindow;

import com.androidtool.lang.util.Utils;
import com.gdubina.tool.langutil.ToolImport;

public class ImportHandler extends ProjectBaseHandler{

	public ImportHandler() {
	}
	
	@Override
	public Object execute(IWorkbenchWindow window, IProject project, String projectName, String projectDir, String workspaceDir) throws ExecutionException {
		FileDialog dialog = new FileDialog(window.getShell(), SWT.OPEN);
		dialog.setFilterNames(new String[] { "Excel File(*.xls)"});
	    dialog.setFilterExtensions(new String[] { "*.xls"}); 
	    dialog.setFilterPath(workspaceDir);
	    dialog.setFileName(projectName + "_strings.xls");
	    String path = dialog.open();
	    System.out.println("Open file: " + path);
	    if(path == null){
	    	return null;
	    }
	    
	    final String folder = Utils.getFileDir(path);
    	final String fileName = Utils.getFileName(path);
    	
    	
	    try {
	    	File fileLog = new File(folder, fileName + "_import.log");
	    	PrintStream printstream = new PrintStream(new FileOutputStream(fileLog), true);
	    	
			ToolImport.run(printstream, projectDir, path);
			
			printstream.close();
			
			project.refreshLocal(IResource.DEPTH_INFINITE, null);
			
	    } catch (IOException e) {
			e.printStackTrace();
		} catch (ParserConfigurationException e) {
			e.printStackTrace();
		} catch (TransformerException e) {
			e.printStackTrace();
		} catch (CoreException e) {
			e.printStackTrace();
		}
		return null;
	}

}
