package com.androidtool.lang.handlers;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.core.resources.IProject;
import org.eclipse.core.resources.ResourcesPlugin;
import org.eclipse.core.runtime.IPath;
import org.eclipse.ui.IWorkbenchWindow;
import org.eclipse.ui.handlers.HandlerUtil;

import com.androidtool.lang.util.Utils;

public abstract class ProjectBaseHandler extends AbstractHandler{

	@Override
	public final Object execute(ExecutionEvent event) throws ExecutionException {
		IWorkbenchWindow window = HandlerUtil.getActiveWorkbenchWindowChecked(event);
		IProject project = Utils.getProjectFromHandle(window);
		if(project == null){
			return null;
		}
		IPath workspace = ResourcesPlugin.getWorkspace().getRoot().getRawLocation();
		
		String workspaceDir = workspace.toOSString();
		String projectDir = workspace.append(project.getFullPath()).toOSString();
		return execute(window, project, project.getName(), projectDir, workspaceDir);
	}
	
	protected abstract Object execute(IWorkbenchWindow window, IProject project, String projectName, String projectDir, String workspaceDir)throws ExecutionException;

}
