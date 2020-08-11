import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Scanner;
import java.util.concurrent.ThreadLocalRandom;

public class Scraper {

    public static void main(String[] args) throws Exception {

        int count = 0;

        String filename = "data/imdb.xlsx";

        String[] columns = {"Title", "Rating", null, "Random Movie", "Rating"};

        ArrayList<Movie> movies = new ArrayList<Movie>();

        System.out.println("\n-------------------------------------------------------------------------------------------");
        System.out.println("IMBd Top Charts to Excel Exporter");
        System.out.println("-------------------------------------------------------------------------------------------\n");

        int numOutput;
        Scanner scanner = new Scanner(System.in);

        do{

            System.out.println("-> Select the number of movies to include (Between 1 and 250)");
            numOutput = scanner.nextInt();

        }while(numOutput < 1 && numOutput > 250);

        final Document document = Jsoup.connect("https://www.imdb.com/chart/top").get();

        for(Element row : document.select("table.chart.full-width tr")) {

            if(count == 0){
                count++;
                continue;
            }

            final String title = row.select(".titleColumn").text();
            final String rating = row.select(".imdbRating").text();

            movies.add(new Movie(title, rating));
            count++;

            if(count == numOutput + 1){
                break;
            }
        }

        // Create Workbook and Sheet
        Workbook workbook = new XSSFWorkbook();
        CreationHelper createHelper = workbook.getCreationHelper();
        Sheet sheet = workbook.createSheet("IMDb Top Charts");

        // Header Styling
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setColor(IndexedColors.BLACK.getIndex());

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        Row headerRow = sheet.createRow(0);

        // Creating Columns
        for(int i = 0; i < columns.length; i++){
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        // Output Movies
        int rowNum = 1;
        for(Movie m : movies){

            Row row = sheet.createRow(rowNum);
            row.createCell(0).setCellValue(m.getTitle());
            row.createCell(1).setCellValue(Double.parseDouble(m.getRating()));

            // Random Movie
            if(rowNum == 1){
                int random = ThreadLocalRandom.current().nextInt(1, numOutput + 1);
                row.createCell(3).setCellValue(movies.get(random).getTitle());
                row.createCell(4).setCellValue(Double.parseDouble(movies.get(random).getRating()));
            }

            rowNum++;
        }

        // Auto Size Columns
        for(int i = 0; i < columns.length; i++){
            sheet.autoSizeColumn(i);
        }

        // File Output
        FileOutputStream fileOut = new FileOutputStream(filename);
        workbook.write((fileOut));
        fileOut.close();
        System.out.println("\n-> IMDb Top " + numOutput + " successfully outputted to imdb-scraper/" + filename);

        // Close Workbook
        workbook.close();

    }
}
