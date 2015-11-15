package org.vaadin;

import com.vaadin.annotations.Theme;
import com.vaadin.annotations.Widgetset;
import com.vaadin.server.VaadinRequest;
import com.vaadin.ui.UI;

@Theme("mytheme")
@SuppressWarnings("serial")
@Widgetset("org.vaadin.AppWidgetSet")
public class MyPortletUI extends UI {

    @Override
    protected void init(VaadinRequest request) {
        setWidth("100%");
        setHeight("800px");
        setContent(new SpreadsheetView());
    }
    
   @Override
   protected void refresh(VaadinRequest request) {
	   super.refresh(request);
	   setContent(new SpreadsheetView());
   }
}
