package com.androidtool.lang.util;

import java.io.File;

import org.eclipse.core.resources.IProject;
import org.eclipse.core.runtime.IAdaptable;
import org.eclipse.jface.viewers.IStructuredSelection;
import org.eclipse.ui.IWorkbenchWindow;

public class Utils {

	public static IProject getProjectFromHandle(IWorkbenchWindow window) {
		IStructuredSelection selection = (IStructuredSelection) window.getSelectionService().getSelection();
        Object firstElement = selection.getFirstElement();
        if (firstElement instanceof IAdaptable){
            return (IProject)((IAdaptable)firstElement).getAdapter(IProject.class);
        }
        return null;
	}
	
	public static String getFileName(String path){
		int fileIndex = path.lastIndexOf(File.separatorChar);
    	int dotIndex = path.lastIndexOf('.');
    	String fileName;
    	if(dotIndex == -1){
    		fileName = path.substring(fileIndex + 1);
    	}else{
    		fileName = path.substring(fileIndex + 1, dotIndex);
    	}
    	return fileName;
	}
	
	public static String getFileDir(String path){
		int fileIndex = path.lastIndexOf(File.separatorChar);
    	return path.substring(0, fileIndex);
	}
}
