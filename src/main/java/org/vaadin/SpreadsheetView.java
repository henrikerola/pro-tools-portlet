package org.vaadin;

import com.liferay.portal.kernel.exception.PortalException;
import com.liferay.portal.kernel.exception.SystemException;
import com.liferay.portal.kernel.repository.model.FileEntry;
import com.liferay.portal.kernel.util.WebKeys;
import com.liferay.portal.theme.ThemeDisplay;
import com.liferay.portlet.documentlibrary.model.DLFolderConstants;
import com.liferay.portlet.documentlibrary.service.DLAppServiceUtil;
import com.vaadin.addon.charts.Chart;
import com.vaadin.addon.charts.model.ChartType;
import com.vaadin.addon.charts.model.Configuration;
import com.vaadin.addon.charts.model.DataSeries;
import com.vaadin.addon.charts.model.DataSeriesItem;
import com.vaadin.addon.spreadsheet.Spreadsheet;
import com.vaadin.event.Action;
import com.vaadin.server.VaadinService;
import com.vaadin.ui.ComboBox;
import com.vaadin.ui.Notification;
import com.vaadin.ui.VerticalLayout;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

public class SpreadsheetView extends VerticalLayout {

    private final Action PLOT_CHART_ACTION = new Action("Plot a chart");

    private Chart chart;
    private Spreadsheet spreadsheet;

    public SpreadsheetView() {
        setSizeFull();
        setSpacing(true);

        ComboBox comboBox = new ComboBox();
        getFileEntries().forEach(fileEntry -> {
            comboBox.addItem(fileEntry);
            comboBox.setItemCaption(fileEntry, fileEntry.getTitle());
        });
        comboBox.addValueChangeListener(e -> {
            setExpandRatio(chart, 0);
            openExcelFileEntry((FileEntry) comboBox.getValue());
        });
        addComponent(comboBox);

        spreadsheet = new Spreadsheet();
        spreadsheet.addActionHandler(new Action.Handler() {
            @Override
            public Action[] getActions(Object target, Object sender) {
                return new Action[]{PLOT_CHART_ACTION};
            }

            @Override
            public void handleAction(Action action, Object sender, Object target) {
                if (action == PLOT_CHART_ACTION) {
                    plotChart();
                    setExpandRatio(chart, 1);
                }
            }
        });
        spreadsheet.setSizeFull();
        addComponent(spreadsheet);
        setExpandRatio(spreadsheet, 1);

        chart = new Chart();
        chart.getConfiguration().setTitle("");
        chart.setSizeFull();
        addComponent(chart);
    }

    private void openExcelFileEntry(FileEntry fileEntry) {
        if (fileEntry == null) {
            spreadsheet.reset();
        } else if (!fileEntry.getExtension().contains("xls")) {
            Notification.show("Not an Excel file");
            spreadsheet.reset();
        } else {
            try {
                spreadsheet.read(fileEntry.getContentStream());
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    private List<FileEntry> getFileEntries() {
        ThemeDisplay themeDisplay = (ThemeDisplay) VaadinService.getCurrentRequest().getAttribute(WebKeys.THEME_DISPLAY);

        long repositoryId = themeDisplay.getScopeGroupId();

        //List<Folder> folders = DLAppServiceUtil.getFolders(repositoryId, DLFolderConstants.DEFAULT_PARENT_FOLDER_ID);
        try {
            return DLAppServiceUtil.getFileEntries(repositoryId, DLFolderConstants.DEFAULT_PARENT_FOLDER_ID).stream()
            		.filter(fileEntry -> fileEntry.getExtension().contains("xls"))
            		.collect(Collectors.toList());
        } catch (PortalException | SystemException e) {
            e.printStackTrace();
        }
        return Collections.emptyList();
    }

    /**
     * Configures the chart component to show a column
     * chart of the selected data in the spreadsheet.
     */
    private void plotChart() {
        // Set the type of chart to a column chart
        Configuration configuration = new Configuration();
        configuration.getChart().setType(ChartType.COLUMN);
        chart.setConfiguration(configuration);

        // Extract the values of all the selected cells and plot them in a chart
        spreadsheet.getCellSelectionManager().getCellRangeAddresses().forEach(this::addDataFromRangeAddress);

        // Force the chart to redraw
        chart.drawChart();
    }

    /**
     * Extracts values from cells in a CellRangeAddress and creates data series
     * for these values and finally adds the data series to the chart.
     *
     * @param selection the CellRangeAddress to extract values from
     */
    private void addDataFromRangeAddress(CellRangeAddress selection) {
        int numRows = selection.getLastRow() - selection.getFirstRow();
        int numCols = selection.getLastColumn() - selection.getFirstColumn();
        Configuration conf = chart.getConfiguration();

        if (numCols > numRows) {
            // We are comparing rows
            chart.getConfiguration().setTitle("Compare rows");
            // Loop through each row and add the data from each cell to a new data series object.
            for (int r = selection.getFirstRow(); r <= selection.getLastRow(); r++) {
                DataSeries series = new DataSeries();
                series.setName("Row " + r);
                Row row = spreadsheet.getActiveSheet().getRow(r);
                if (row != null) {
                    for (int c = selection.getFirstColumn(); c <= selection.getLastColumn(); c++) {
                        Cell cell = row.getCell(c);
                        // If the cell is a numeric value, add the numeric value, otherwise add null instead
                        if (cell != null && cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            series.add(new DataSeriesItem(CellReference.convertNumToColString(c),
                                    cell.getNumericCellValue()));
                        } else {
                            series.add(new DataSeriesItem(CellReference.convertNumToColString(c), null));
                        }
                    }
                }
                // Add the data series to the chart
                conf.addSeries(series);
            }
        } else {
            // We are comparing columns
            chart.getConfiguration().setTitle("Compare columns");
            // Loop through each column and add the data from each cell to a new data series object.
            for (int c = selection.getFirstColumn(); c <= selection.getLastColumn(); c++) {
                DataSeries series = new DataSeries();
                series.setName("Col " + CellReference.convertNumToColString(c));
                for (int r = selection.getFirstRow(); r <= selection.getLastRow(); r++) {
                    Row row = spreadsheet.getActiveSheet().getRow(r);
                    if (row != null) {
                        Cell cell = row.getCell(c);
                        // If the cell is a numeric value, add the numeric value, otherwise add null instead
                        if (cell != null && cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                            series.add(new DataSeriesItem(row.getRowNum() + 1, cell.getNumericCellValue()));
                        } else {
                            series.add(new DataSeriesItem(row.getRowNum() + 1, null));
                        }
                    }
                }
                // Add the data series to the chart
                conf.addSeries(series);
            }
        }
    }
}
