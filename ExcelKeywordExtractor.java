package com.example;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.icepear.echarts.Bar;
import org.icepear.echarts.Line;
import org.icepear.echarts.Pie;
import org.icepear.echarts.Scatter;
import org.icepear.echarts.render.Engine;
//import org.icepear.echarts.utils.ImageUtils;
import java.io.*;
import java.util.*;
import java.util.regex.*;
public class ExcelKeywordExtractor {
    // Apache help from https://howtodoinjava.com/java/library/readingwriting-excel-files-in-java-poi-tutorial/
    // Get value from cache help from https://www.baeldung.com/apache-poi-read-cell-value-formula
    // Maven help from https://maven.apache.org/guides/getting-started/maven-in-five-minutes.html
    // Apache.com for the dependencies I needed to add
    //E graph help from https://github.com/ECharts-Java/ECharts-Java
    // Help from Jason Paarlberg for using Apche, and Maven and the idea to process it that way and get the value from cache.
    private static List<String> keywords = new ArrayList<>();
    private static String regGraph;
    private static String indFilePath;
    private static String depFilePath;
    private static String regressionType;

    private static String sort;
    private static Map<String, Integer> yearsIndex = new HashMap<>();

    private static Map<String, Map<String, Integer>> yearsMap = new HashMap<String, Map<String, Integer>>();
    private static Map<String, Integer> assIndex = new HashMap<>();
    private static boolean summary = false;
    private static Map<String, Map<String, Double>> associationYearValueMap = new HashMap<>();
    private static Map<String, Integer> associationRowMap = new HashMap<>();
    private static int assRowIndex = 0;

    private static Map<List<String>, List<String>> CombineMap = new HashMap<>();

    public static void main(String AssPath, Vector<Integer> years, String depFile, String indFile, String regType, boolean sum, String graphType, String sortType) {
        sort = sortType;
        summary = sum;
        regGraph = graphType;
        System.out.println("Graph Type: " + regGraph);
        indFilePath = indFile;
        depFilePath = depFile;
        regressionType = regType;
//        String assPath = "Association";
//        String assFile = "Associations.txt";
        // Load association keywords from user's file
        File wordFile = new File(AssPath);
        if (wordFile.exists()) {
            try (BufferedReader reader = new BufferedReader(new FileReader(wordFile))) {
                String line;
                while ((line = reader.readLine()) != null) {
                    if (line.contains("=")){
                        String[] parts = line.split("=");
                        if (CombineMap.containsKey(List.of(parts[0].trim().toUpperCase()))) {
                            CombineMap.get(List.of(parts[0].trim().toUpperCase())).add(parts[1].trim().toUpperCase());
                        } else {
                            CombineMap.put(List.of(parts[0].trim().toUpperCase()), new ArrayList<>(List.of(parts[1].trim().toUpperCase())));
                        }
                        keywords.add(parts[1].trim().toUpperCase());
                        System.out.println("CombineMap: " + parts[0].trim().toUpperCase() + " = " + parts[1].trim().toUpperCase());
                        System.out.println();
                        continue;

                    }
                    keywords.add(line.trim().toUpperCase());
                    System.out.println("keyword: " + line);  //
                }
                System.out.println(CombineMap + "CombineMap");
            } catch (IOException e) {
                System.out.println("Error reading Associations.txt file");
                e.printStackTrace();
            }
        }
        if (keywords.isEmpty()) {
            System.out.println("No keywords found, exiting.");
            return; // No keywords found, so nothing would be added
        }
        Scanner scanner = new Scanner(System.in);
        // Prompt for Excel files
        //System.out.print("Enter comma-separated Excel file paths (e.g., Excel/file1.xlsx,Excel/file2.xlsm): ");
        String[] filePaths = depFilePath.trim().split(",");
        List<File> selectedExcelFiles = new ArrayList<>();
        for (String path : filePaths) {
            File file = new File(path.trim());
            if (file.exists() && (path.endsWith(".xlsx") || path.endsWith(".xlsm"))) {
                selectedExcelFiles.add(file);
            } else {
                System.out.println("Skipping invalid or missing file: " + path);
            }
        }
        if (selectedExcelFiles.isEmpty()) {
            System.out.println("No valid Excel files provided. Exiting.");
            return;
        }
        // Prompt for years
        //System.out.print(" years to extract (e.g., 2008 2009 2010): ");
        //String[] yearTokens = scanner.nextLine().trim().split("\\s+");
        Set<String> validYears = new HashSet<>(years.stream().map(String::valueOf).toList());
        System.out.println("Valid years: " + validYears);

        // Prompt for output file
        System.out.print("Enter name for output file (e.g., output): ");
        String outputFileName = scanner.nextLine().trim();
        try (Workbook outputWorkbook = new XSSFWorkbook()) {
            Sheet outputSheet = outputWorkbook.createSheet("CombinedOutput");
            Map<CellStyle, CellStyle> styleMap = new HashMap<>();
            int currentRowIndex = 0;
            for (File file : selectedExcelFiles) {
                try (FileInputStream fis = new FileInputStream(file);
                     Workbook inputWorkbook = WorkbookFactory.create(fis)) {
                    for (int s = 0; s < inputWorkbook.getNumberOfSheets(); s++) {
                        Sheet inputSheet = inputWorkbook.getSheetAt(s);
                        // Skip sheets without any valid keyword matches
                        if (!sheetContainsKeyword(inputSheet, keywords)) continue;
                        // Copy rows based on valid years or keywords so that is only has those
                        for (int rowIndex = 0; rowIndex <= inputSheet.getLastRowNum(); rowIndex++) {
                            Row srcRow = inputSheet.getRow(rowIndex);
                            if (srcRow == null) continue;
                            assRowIndex = currentRowIndex;// Skip empty rows
                            // Check if the row matches a valid year or contains a keyword
                            if (isRowMatching(srcRow, validYears)) {
                                Row destRow = outputSheet.createRow(currentRowIndex++);
                                // Copy the row exactly as it is from the source
                                for (int col = 0; col < srcRow.getLastCellNum(); col++) {

                                    Cell srcCell = srcRow.getCell(col);
                                    Cell destCell = destRow.createCell(col);
                                    if (srcCell != null) {

                                        copyCell(srcCell, destCell, outputWorkbook, styleMap);
                                    } else {
                                        destCell.setBlank();
                                    }
                                }
                            }
                        }
                        currentRowIndex++; // Add a buffer row
                    }
                } catch (Exception e) {
                    System.out.println("Error reading file: " + file.getName());
                    e.printStackTrace();
                }
            }

            // Print graph of data here
            // addGraph();
            Row destRow = outputSheet.createRow(currentRowIndex++);
            addGraph(destRow, outputWorkbook, currentRowIndex);
            if (currentRowIndex == 0) {
                System.out.println("No matching data found.");
            } else {
                try (FileOutputStream fos = new FileOutputStream(outputFileName + ".xlsx")) {
                    outputWorkbook.write(fos);
                    System.out.println("Output saved to " + outputFileName + ".xlsx with " + currentRowIndex + " rows.");
                }
            }
        } catch (IOException e) {
            System.out.println("Error creating output file");
            e.printStackTrace();
        }

    }
    public static double rSquared(double[] x, double[] y) {
        int len = Math.min(x.length, y.length);

        double[] xTrimmed = new double[len];
        double[] yTrimmed = new double[len];

        // Copy from the end: keep the most recent values
        System.arraycopy(x, x.length - len, xTrimmed, 0, len);
        System.arraycopy(y, y.length - len, yTrimmed, 0, len);

        double meanY = 0.0;
        for (double val : yTrimmed) {
            meanY += val;
        }
        meanY /= len;

        double ssTot = 0.0;
        double ssRes = 0.0;
        for (int i = 0; i < len; i++) {
            double yPred = xTrimmed[i]; // Replace with predicted value if modeling
            ssRes += Math.pow(yTrimmed[i] - yPred, 2);
            ssTot += Math.pow(yTrimmed[i] - meanY, 2);
        }

        return 1 - (ssRes / ssTot);
    }

    public static void addGraph (Row destRow, Workbook outputWorkbook, int currentRowIndex) {

        // int rowIndexShift = 0;
        for (Map.Entry<String, Integer> assEntry : assIndex.entrySet()) {
            String association = assEntry.getKey();
            //System.out.println(assEntry.getValue() + "adsasdasdsa");

            int rowIndex = associationRowMap.getOrDefault(association, -1); // Default to -1 if not found

            // Need to create a map with each association and the row it is in the dest sheet, so that we can pull the values from there
            //
           //
            //System.out.println("Association: " + association + ", Row Index: " + rowIndex);
           // System.out.println("Association Row Map: " + associationRowMap);
            Map<String, Double> yearValueMap = new HashMap<>();
            //System.out.println(yearsIndex);
            for (Map.Entry<String, Integer> yearEntry : yearsIndex.entrySet()) {
                //System.out.println("Association: " + association + ", Year: " + yearEntry.getKey());
                String year = yearEntry.getKey();
                int columnIndex = yearEntry.getValue();
                Row row = destRow.getSheet().getRow(rowIndex - 1);
                //rowIndexShift++;
                if (row != null) {
                    Cell cell = row.getCell(columnIndex);
                    //System.out.println(cell);
                   // System.out.println("Association: " + association + ", Year: " + year + ", Column Index: " + columnIndex);
                    if (cell.getCellType() == CellType.NUMERIC) {
                        //System.out.println("Association: " + association + ", Year: " + year + ", Value: " + cell.getNumericCellValue());
                        yearValueMap.put(year, cell.getNumericCellValue());
                    }
                }
            }
            associationYearValueMap.put(association, yearValueMap);
        }
        System.out.println("Association-Year-Value Map: " + associationYearValueMap);
        System.out.println("Year Map: " + yearsMap);

        // Create a new sheet for the graph
        // This is the eGraph I was trying to do it in both I feel like it is a sick feature

        int graphIndexCombine = 0;

// Create a single Line chart for all associations

        //System.out.println("CombineMap: " + CombineMap);

        for (List<String> key : CombineMap.keySet()) {
            System.out.println("CombineMap: " + key);
            System.out.println("CombineMap: " + CombineMap.get(key));
            List<String> sortedYears2 = new ArrayList<>(yearsIndex.keySet());
            Collections.sort(sortedYears2);
            Line combinedLineChart = new Line()
                    .setLegend()
                    .setTooltip("item")
                    .addXAxis(sortedYears2.toArray(new String[0])) // Use sorted years as the X-axis
                    .addYAxis();
            System.out.println(key);

            // System.out.println(key);
            // System.out.println("CombineMap: " + key + " = " + CombineMap.get(key));
            for (Map.Entry<String, Map<String, Double>> entry : associationYearValueMap.entrySet()) {
                String association = entry.getKey();

                //System.out.println(CombineMap + " " + CombineMap.keySet());
                //System.out.println();
                if (!CombineMap.get(key).contains(association)) {
                    continue;
                }
                //if (!association.toLowerCase().contains("reported revenue")) continue;
                Map<String, Double> yearValueMap = entry.getValue();

                boolean hasData = yearValueMap.values().stream().anyMatch(value -> value != 0);
                if (!hasData) {
                    System.out.println("Skipping data for association: " + association + " as it has no data.");
                    continue; // Skip this association if no data is present
                }

                List<Map.Entry<String, Double>> sortedYearValueList = new ArrayList<>(yearValueMap.entrySet());
                sortedYearValueList.sort(Map.Entry.comparingByKey());
                List<Number> valuesForYears = new ArrayList<>();
                //System.out.println(sortedYears2);
                for (String year : sortedYears2) {
                    //System.out.println("yearrrr" + year);
                    valuesForYears.add(yearValueMap.getOrDefault(year, 0.0)); // Fill in 0.0 if missing
                }
                //System.out.println("Association: " + association + ", Values: " + valuesForYears);

                combinedLineChart.addSeries(association, valuesForYears.toArray(new Number[0]));

            }


            Engine enginee = new Engine();
            System.out.println("Rendering combined graph...");


// Render the combined graph
            enginee.render("combined_graph" + graphIndexCombine + ".html", combinedLineChart);

            try {
                System.out.println("Converting HTML to PNG...");
                // Specify full paths for HTML and PNG
                String htmlFilePath = new File("combined_graph" + graphIndexCombine + ".html").getAbsolutePath();
                String pngFilePath = new File("combined_chart_output" + graphIndexCombine + ".png").getAbsolutePath();
                String command = "node convertHtmlToPng.js " + htmlFilePath + " " + pngFilePath;
                Process process = Runtime.getRuntime().exec(command);
                process.waitFor();
                System.out.println("Combined chart saved as " + pngFilePath);

                // Now add the image to the Excel file
                Sheet imageSheet = outputWorkbook.createSheet("CombinedGraph" + graphIndexCombine);
                // Make sure the sheet is properly sized
                imageSheet.setColumnWidth(0, 5000);  // Set column width
                imageSheet.setColumnWidth(1, 5000);
                // Load the PNG file as a byte array
                InputStream inputStream = new FileInputStream(pngFilePath);
                byte[] imageBytes = inputStream.readAllBytes();
                int pictureIdx = outputWorkbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
                inputStream.close();
                // Create a drawing patriarch (responsible for managing images in Excel)
                Drawing<?> drawing = imageSheet.createDrawingPatriarch();
                // anchor sets where it goes
                ClientAnchor anchor = outputWorkbook.getCreationHelper().createClientAnchor();
                anchor.setCol1(0);
                anchor.setRow1(0);
                anchor.setCol2(10);
                anchor.setRow2(40);
                // Create the picture in the Excel sheet using the anchor
                drawing.createPicture(anchor, pictureIdx);
                System.out.println("Combined image inserted into Excel file.");
            } catch (Exception e) {
                e.printStackTrace();
                System.out.println("Error converting HTML to PNG.");
            }
            graphIndexCombine++;
        }


        if (regGraph.equals("bar")) {
            int graphIndex = 0;
        for (Map.Entry<String, Map<String, Double>> entry : associationYearValueMap.entrySet()) {

            Bar bar = new Bar()
                    .setLegend()
                    .setTooltip("item")

                    .addXAxis(keywords.toArray(new String[0]))
                    .addYAxis();
            String association = entry.getKey();
            Map<String, Double> yearValueMap = entry.getValue();

            boolean hasData = yearValueMap.values().stream().anyMatch(value -> value != 0);
            if (!hasData) {
                System.out.println("Skipping graph for association: " + association + " as it has no data.");
                continue; // Skip this graph if no data is present
            }
            for (Map.Entry<String, Double> yearEntry : yearValueMap.entrySet()) {


                String year = yearEntry.getKey();
                Double value = yearEntry.getValue();

                if (value != 0) { // Add only if the value is not 0
                    bar.addSeries(association + " (" + year + ")", new Number[] { value });

                }

            }
            Engine engine = new Engine();
            // The render method will generate our EChart into a HTML file saved locally in the current directory.
            // The name of the HTML can also be set by the first parameter of the function.
            System.out.println(regGraph);

            engine.render("index" + graphIndex + ".html", bar);
            try {
                System.out.println("Converting HTML to PNG...");
                // Specify full paths for HTML and PNG
                String htmlFilePath = new File("index" + graphIndex + ".html").getAbsolutePath();
                String pngFilePath = new File("chart_output.png").getAbsolutePath();
                String command = "node convertHtmlToPng.js " + htmlFilePath + " " + pngFilePath;
                Process process = Runtime.getRuntime().exec(command);
                process.waitFor();
                System.out.println("Chart saved as " + pngFilePath);
                // Now add the image to the Excel file
                Sheet imageSheet = outputWorkbook.createSheet("Graph" + graphIndex);
                graphIndex++;
                // Make sure the sheet is properly sized
                imageSheet.setColumnWidth(0, 5000);  // Set column width
                imageSheet.setColumnWidth(1, 5000);
                // Load the PNG file as a byte array
                InputStream inputStream = new FileInputStream(pngFilePath);
                byte[] imageBytes = inputStream.readAllBytes();
                int pictureIdx = outputWorkbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
                inputStream.close();
                // Create a drawing patriarch (responsible for managing images in Excel)
                Drawing<?> drawing = imageSheet.createDrawingPatriarch();
                // anchor sets where it goes
                ClientAnchor anchor = outputWorkbook.getCreationHelper().createClientAnchor();
                anchor.setCol1(0);
                anchor.setRow1(0);
                anchor.setCol2(10);
                anchor.setRow2(40);
                // Create the picture in the Excel sheet using the anchor
                drawing.createPicture(anchor, pictureIdx);
                System.out.println("Image inserted into Excel file.");
            } catch (Exception e) {
                e.printStackTrace();
                System.out.println("Error converting HTML to PNG.");
            }
        }} else if (regGraph.equals("line")) {
            System.out.println(regGraph);
            int graphIndex = 0;
            List<String> sortedYears = new ArrayList<>(yearsIndex.keySet());
            Collections.sort(sortedYears);  // Sort the years properly

            for (Map.Entry<String, Map<String, Double>> entry : associationYearValueMap.entrySet()) {
                Line line = new Line()
                        .setLegend()
                        .setTooltip("item")
                        .addXAxis(sortedYears.toArray(new String[0])) // X axis is years
                        .addYAxis();

                String association = entry.getKey();
                Map<String, Double> yearValueMap = entry.getValue();

                boolean hasData = yearValueMap.values().stream().anyMatch(value -> value != 0);
                if (!hasData) {
                    System.out.println("Skipping graph for association: " + association + " as it has no data.");
                    continue;
                }

                List<Number> valuesForYears = new ArrayList<>();
                for (String year : sortedYears) {
                    valuesForYears.add(yearValueMap.getOrDefault(year, 0.0)); // Fill in 0.0 if missing
                }

                line.addSeries(association, valuesForYears.toArray(new Number[0]));

// Calculate and add a trendline
                double[] trendline = calculateTrendline(valuesForYears);
                Number[] trendlineNumbers = Arrays.stream(trendline).boxed().toArray(Number[]::new);
                line.addSeries(association + " Trendline", trendlineNumbers);


                Engine engine = new Engine();
                System.out.println("Rendering graph for: " + association);
                engine.render("index" + graphIndex + ".html", line);

                try {
                    System.out.println("Converting HTML to PNG...");
                    String htmlFilePath = new File("index" + graphIndex + ".html").getAbsolutePath();
                    String pngFilePath = new File("chart_output" + graphIndex + ".png").getAbsolutePath();
                    String command = "node convertHtmlToPng.js " + htmlFilePath + " " + pngFilePath;
                    Process process = Runtime.getRuntime().exec(command);
                    process.waitFor();
                    System.out.println("Chart saved as " + pngFilePath);

                    // Insert into Excel
                    Sheet imageSheet = outputWorkbook.createSheet("Graph" + graphIndex);
                    graphIndex++;
                    imageSheet.setColumnWidth(0, 5000);
                    imageSheet.setColumnWidth(1, 5000);

                    InputStream inputStream = new FileInputStream(pngFilePath);
                    byte[] imageBytes = inputStream.readAllBytes();
                    int pictureIdx = outputWorkbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
                    inputStream.close();

                    Drawing<?> drawing = imageSheet.createDrawingPatriarch();
                    ClientAnchor anchor = outputWorkbook.getCreationHelper().createClientAnchor();
                    anchor.setCol1(0);
                    anchor.setRow1(0);
                    anchor.setCol2(10);
                    anchor.setRow2(40);
                    drawing.createPicture(anchor, pictureIdx);

                    System.out.println("Image inserted into Excel file.");
                } catch (Exception e) {
                    e.printStackTrace();
                    System.out.println("Error converting HTML to PNG.");
                }
            }}else if (regGraph.equals("pie")) {
            System.out.println(regGraph);
            int graphIndex = 0;
            List<String> sortedYears = new ArrayList<>(yearsIndex.keySet());
            Collections.sort(sortedYears);  // Sort the years properly
            int yearsindex = 0;
            for (Map.Entry<String, Map<String, Double>> entry : associationYearValueMap.entrySet()) {
                Pie pie = new Pie()
                        .setLegend()
                        .setTooltip("item");
                        //.addXAxis(sortedYears.toArray(new String[0])) // X axis is years
                        //.addYAxis();

                String association = entry.getKey();
                Map<String, Double> yearValueMap = entry.getValue();

                boolean hasData = yearValueMap.values().stream().anyMatch(value -> value != 0);
                if (!hasData) {
                    System.out.println("Skipping graph for association: " + association + " as it has no data.");
                    continue;
                }
                List<Number> valuesForYears = new ArrayList<>();

                for (String year : sortedYears) {

                    valuesForYears.add(yearValueMap.getOrDefault(year, 0.0)); // Fill in 0.0 if missing
                }
                //String year = sortedYears.get(yearsindex);
                pie.addSeries(association, valuesForYears.toArray(new Number[0]));
               // yearsindex++;


                Engine engine = new Engine();
                System.out.println("Rendering graph for: " + association);
                engine.render("index" + graphIndex + ".html", pie);

                try {
                    System.out.println("Converting HTML to PNG...");
                    String htmlFilePath = new File("index" + graphIndex + ".html").getAbsolutePath();
                    String pngFilePath = new File("chart_output" + graphIndex + ".png").getAbsolutePath();
                    String command = "node convertHtmlToPng.js " + htmlFilePath + " " + pngFilePath;
                    Process process = Runtime.getRuntime().exec(command);
                    process.waitFor();
                    System.out.println("Chart saved as " + pngFilePath);

                    // Insert into Excel
                    Sheet imageSheet = outputWorkbook.createSheet("Graph" + graphIndex);
                    graphIndex++;
                    imageSheet.setColumnWidth(0, 5000);
                    imageSheet.setColumnWidth(1, 5000);

                    InputStream inputStream = new FileInputStream(pngFilePath);
                    byte[] imageBytes = inputStream.readAllBytes();
                    int pictureIdx = outputWorkbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
                    inputStream.close();

                    Drawing<?> drawing = imageSheet.createDrawingPatriarch();
                    ClientAnchor anchor = outputWorkbook.getCreationHelper().createClientAnchor();
                    anchor.setCol1(0);
                    anchor.setRow1(0);
                    anchor.setCol2(10);
                    anchor.setRow2(40);
                    drawing.createPicture(anchor, pictureIdx);

                    System.out.println("Image inserted into Excel file.");
                } catch (Exception e) {
                    e.printStackTrace();
                    System.out.println("Error converting HTML to PNG.");
                }
            }}else if (regGraph.equals("scatter")) {
            System.out.println(regGraph);
            int graphIndex = 0;
            List<String> sortedYears = new ArrayList<>(yearsIndex.keySet());
            Collections.sort(sortedYears);  // Sort the years properly

            for (Map.Entry<String, Map<String, Double>> entry : associationYearValueMap.entrySet()) {

                Scatter scatter = new Scatter()
                        .setLegend()
                        .setTooltip("item")
                        .addXAxis(sortedYears.toArray(new String[0])) // X axis is years
                        .addYAxis();

                String association = entry.getKey();
                Map<String, Double> yearValueMap = entry.getValue();

                boolean hasData = yearValueMap.values().stream().anyMatch(value -> value != 0);
                if (!hasData) {
                    System.out.println("Skipping graph for association: " + association + " as it has no data.");
                    continue;
                }

                List<Number> valuesForYears = new ArrayList<>();
                for (String year : sortedYears) {
                    valuesForYears.add(yearValueMap.getOrDefault(year, 0.0)); // Fill in 0.0 if missing
                }

                scatter.addSeries(association, valuesForYears.toArray(new Number[0]));

                // Calculate and add a trendline
                double[] trendline = calculateTrendline(valuesForYears);
                Number[] trendlineNumbers = Arrays.stream(trendline).boxed().toArray(Number[]::new);
                scatter.addSeries(association + " Trendline", trendlineNumbers);

                Engine engine = new Engine();
                System.out.println("Rendering graph for: " + association);
                engine.render("index" + graphIndex + ".html", scatter);

                try {
                    System.out.println("Converting HTML to PNG...");
                    String htmlFilePath = new File("index" + graphIndex + ".html").getAbsolutePath();
                    String pngFilePath = new File("chart_output" + graphIndex + ".png").getAbsolutePath();
                    String command = "node convertHtmlToPng.js " + htmlFilePath + " " + pngFilePath;
                    Process process = Runtime.getRuntime().exec(command);
                    process.waitFor();
                    System.out.println("Chart saved as " + pngFilePath);

                    // Insert into Excel
                    Sheet imageSheet = outputWorkbook.createSheet("Graph" + graphIndex);
                    graphIndex++;
                    imageSheet.setColumnWidth(0, 5000);
                    imageSheet.setColumnWidth(1, 5000);

                    InputStream inputStream = new FileInputStream(pngFilePath);
                    byte[] imageBytes = inputStream.readAllBytes();
                    int pictureIdx = outputWorkbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
                    inputStream.close();

                    Drawing<?> drawing = imageSheet.createDrawingPatriarch();
                    ClientAnchor anchor = outputWorkbook.getCreationHelper().createClientAnchor();
                    anchor.setCol1(0);
                    anchor.setRow1(0);
                    anchor.setCol2(10);
                    anchor.setRow2(40);
                    drawing.createPicture(anchor, pictureIdx);

                    System.out.println("Image inserted into Excel file.");

                // Regression Lines here
                    // anchor.setCol1(12);
                    // anchor.setRow1(0);
                    // anchor.setCol2(22);
                    // anchor.setRow2(40);


                } catch (Exception e) {
                    e.printStackTrace();
                    System.out.println("Error converting HTML to PNG.");
                }
            }}
        if (summary) {
            System.out.println(sort);

            //set sort type to low to high if nothing is specified
            if (sort == null || sort.isEmpty()) {
                sort = "LowToHigh";
            }

            Sheet summarySheet = outputWorkbook.createSheet("Summary");

            List<String> allYears = new ArrayList<>();
            List<String> allAssociations = new ArrayList<>();

            for (Map.Entry<String, Map<String, Double>> entry : associationYearValueMap.entrySet()) {
                String association = entry.getKey();
                if (!allAssociations.contains(association)) {
                    allAssociations.add(association);
                }

                for (String year : entry.getValue().keySet()) {
                    if (!allYears.contains(year)) {
                        allYears.add(year);
                    }
                }
            }

            Collections.sort(allAssociations);
            Collections.sort(allYears);

            if (sort.equals("LowToHigh")) {

            } else if (sort.equals("HighToLow")) {//
                Collections.reverse(allYears);
            }


            else if (sort.equals("Alpha")) {
                Collections.sort(allAssociations, String.CASE_INSENSITIVE_ORDER);
            } else if (sort.equals("BackAlpha")) {
                Collections.reverse(allAssociations);
            } else {
                Collections.sort(allYears);
            }
//            else {
//                System.out.println("Invalid sort type. Defaulting to LowToHigh.");
//            }

// Create header row
            Row headerRow = summarySheet.createRow(0);
            headerRow.createCell(0).setCellValue("Year");
            int colIndex = 1;
            Map<String, Integer> associationColumnIndex = new HashMap<>();

            for (String association : allAssociations) {
                headerRow.createCell(colIndex).setCellValue(association);
                headerRow.getSheet().autoSizeColumn(colIndex);
                associationColumnIndex.put(association, colIndex);
                colIndex++;
            }

// Fill in data rows
            int rowIndex = 1;
            for (String year : allYears) {
                Row row = summarySheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(year);

                for (String association : allAssociations) {
                    Map<String, Double> yearValueMap = associationYearValueMap.get(association);
                    if (yearValueMap != null && yearValueMap.containsKey(year)) {
                        int assocCol = associationColumnIndex.get(association);
                        row.createCell(assocCol).setCellValue(yearValueMap.get(year));
                    }
                }
            }
        }




    }
    private static double[] calculateTrendline(List<Number> values) {
        int n = values.size();
        if (n == 0) {
            throw new IllegalArgumentException("Values list cannot be empty.");
        }

        double sumX = 0;
        double sumY = 0;
        double sumXY = 0;
        double sumX2 = 0;

        for (int i = 0; i < n; i++) {
            double x = i; // X values are indices
            double y = values.get(i).doubleValue();

            sumX += x;
            sumY += y;
            sumXY += x * y;
            sumX2 += x * x;
        }

        double slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
        double intercept = (sumY - slope * sumX) / n;

        double[] trendline = new double[n];
        for (int i = 0; i < n; i++) {
            trendline[i] = slope * i + intercept;
        }

        return trendline;
    }

        // Check if the sheet contains any of the association keywords
    private static boolean sheetContainsKeyword(Sheet sheet, List<String> keywords) {
        boolean found = false;
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    String val = cell.getStringCellValue().toUpperCase();
                    for (String keyword : keywords) {
                        if (val.contains(keyword)) {
                            found = true;
                            assIndex.put(keyword, row.getRowNum());// Add keyword and column index to yearsIndex
                           // System.out.println("Keyword found: " + keyword + " in row index: " + row.getRowNum());
                        }// Set found to true if any keyword is found
                    }
                }
            }
        }
        //System.out.println(assIndex);
        return found;
    }

    // Check if the row contains valid years or any keywords
    private static boolean isRowMatching(Row row, Set<String> validYears) {
        Pattern yearPattern = Pattern.compile("\\b\\d{4}\\b");
        boolean yearFound = false;
        String matchedKeyword = null;

        for (Cell cell : row) {
            if (cell != null) {
                String cellValue = cell.toString().toUpperCase();

                // Check if this cell matches any keyword
                for (String keyword : keywords) {
                    if (cellValue.equals(keyword)) {
                        matchedKeyword = keyword;
                        associationRowMap.put(keyword, assRowIndex + 1);
                        assIndex.put(keyword, row.getRowNum());
                        break;
                    }
                }

                // If the cell has a valid year
                Matcher yearMatcher = yearPattern.matcher(cellValue);
                if (yearMatcher.find()) {
                    String foundYear = yearMatcher.group();
                    if (validYears.contains(foundYear)) {
                        yearsIndex.put(foundYear, cell.getColumnIndex());
                        //System.out.println("Year found: " + foundYear + " in column index: " + cell.getColumnIndex());
                        yearFound = true;

                        // If a keyword was matched earlier in this row, update yearsMap for that keyword
                        if (matchedKeyword != null) {
                            yearsMap.computeIfAbsent(matchedKeyword, k -> new HashMap<>())
                                    .put(foundYear, cell.getColumnIndex());
                        }
                    }
                }
            }
        }

        return yearFound || matchedKeyword != null;
    }




    // Check if the row is blank
    private static boolean isRowBlank(Row row) {
        if (row == null) return true;
        for (int i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null && cell.getCellType() != CellType.BLANK && !cell.toString().trim().isEmpty()) {
                return false;
            }
        }
        return true;
    }
    // Copy a cell's value and style to a new cell through the cache since I can't add macros
    private static void copyCell(Cell oldCell, Cell newCell, Workbook outputWorkbook, Map<CellStyle, CellStyle> styleMap) {
        CellStyle newStyle = oldCell.getCellStyle();
        if (oldCell.getSheet().getWorkbook() != outputWorkbook) {
            if (!styleMap.containsKey(newStyle)) {
                CellStyle clonedStyle = outputWorkbook.createCellStyle();
                clonedStyle.cloneStyleFrom(newStyle);
                styleMap.put(newStyle, clonedStyle);
                newStyle = clonedStyle;
            } else {
                newStyle = styleMap.get(newStyle);
            }
        }
        newCell.setCellStyle(newStyle);

        CellType type = oldCell.getCellType();
        if (type == CellType.FORMULA) type = oldCell.getCachedFormulaResultType();
        // this determines what type the data is
        switch (type) {
            case STRING -> newCell.setCellValue(oldCell.getStringCellValue());
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(oldCell)) {
                    newCell.setCellValue(oldCell.getDateCellValue());
                } else {
                    newCell.setCellValue(oldCell.getNumericCellValue());
                }
            }
            case BOOLEAN -> newCell.setCellValue(oldCell.getBooleanCellValue());
            case BLANK -> newCell.setBlank();
            default -> newCell.setCellValue(oldCell.toString());
        }
    }
}
