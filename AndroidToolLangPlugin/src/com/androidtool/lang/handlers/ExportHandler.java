package com.androidtool.lang.handlers;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;

import javax.xml.parsers.ParserConfigurationException;

import org.eclipse.core.commands.ExecutionException;
import org.eclipse.core.resources.IProject;
import org.eclipse.swt.SWT;
import org.eclipse.swt.program.Program;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.ui.IWorkbenchWindow;
import org.xml.sax.SAXException;

import com.androidtool.lang.util.Utils;
import com.gdubina.tool.langutil.ToolExport;

/**
 * Our sample handler extends AbstractHandler, an IHandler base class.
 * @see org.eclipse.core.commands.IHandler
 * @see org.eclipse.core.commands.AbstractHandler
 */
public class ExportHandler extends ProjectBaseHandler {
	/**
	 * The constructor.
	 */
	public ExportHandler() {
	}

	/**
	 * the command has been executed, so extract extract the needed information
	 * from the application context.
	 */
	@Override
	public Object execute(IWorkbenchWindow window, IProject project, String projectName, String projectDir, String workspaceDir) throws ExecutionException {
		
		FileDialog dialog = new FileDialog(window.getShell(), SWT.SAVE);
	    dialog.setFilterNames(new String[] { "Excel File(*.xls)"});
	    dialog.setFilterExtensions(new String[] { "*.xls"}); 
	    dialog.setFilterPath(workspaceDir);
	    dialog.setFileName(projectName + "_strings.xls");
	    String path = dialog.open();
	    
	    System.out.println("Save to: " + path);
	    if(path == null){
	    	return null;
	    }
	    
	    final String folder = Utils.getFileDir(path);
	    final String fileName = Utils.getFileName(path);
	    
	    try {
	    	
	    	File fileLog = new File(folder, fileName + "_export.log");
	    	PrintStream printstream = new PrintStream(new FileOutputStream(fileLog), true);
	    	
			ToolExport.run(printstream, projectDir, path);

			printstream.close();
			
			Program.launch(folder);
			Program.launch(fileLog.getAbsolutePath());
			
		} catch (SAXException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (ParserConfigurationException e) {
			e.printStackTrace();
		}
		return null;
	}
}
