package me.aboullaite.corenlp.sentimentanalysis.services;

import edu.stanford.nlp.ling.CoreAnnotations;
import edu.stanford.nlp.neural.rnn.RNNCoreAnnotations;
import edu.stanford.nlp.pipeline.Annotation;
import edu.stanford.nlp.pipeline.StanfordCoreNLP;
import edu.stanford.nlp.sentiment.SentimentCoreAnnotations;
import edu.stanford.nlp.trees.Tree;
import edu.stanford.nlp.util.CoreMap;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.*;

@Service
public class SentimentAnalyzerService {

    private static int analyse(String tweet) {

        Properties props = new Properties();
        props.setProperty("annotators", "tokenize, ssplit, pos, parse, sentiment");
        StanfordCoreNLP pipeline = new StanfordCoreNLP(props);
        Annotation annotation = pipeline.process(tweet);
        for (CoreMap sentence : annotation.get(CoreAnnotations.SentencesAnnotation.class)) {
            Tree tree = sentence.get(SentimentCoreAnnotations.SentimentAnnotatedTree.class);
            return RNNCoreAnnotations.getPredictedClass(tree);
        }
        return 0;
    }


    public void findSentiment(String line) {

        Properties props = new Properties();

        props.setProperty("annotators", "tokenize, ssplit, parse, sentiment");

        StanfordCoreNLP pipeline = new StanfordCoreNLP(props);

        int mainSentiment = 0;

        if (line != null && line.length() > 0) {

            int longest = 0;

            Annotation annotation = pipeline.process(line);

            for (CoreMap sentence : annotation.get(CoreAnnotations.SentencesAnnotation.class)) {

                Tree tree = sentence.get(SentimentCoreAnnotations.SentimentAnnotatedTree.class);

                int sentiment = RNNCoreAnnotations.getPredictedClass(tree);

                String partText = sentence.toString();

                if (partText.length() > longest) {

                    mainSentiment = sentiment;

                    longest = partText.length();

                }


            }

        }

        if (mainSentiment == 2 || mainSentiment > 4 || mainSentiment < 0) {

            return;

        }

        //TweetWithSentiment tweetWithSentiment = new TweetWithSentiment(line, toCss(mainSentiment));
        return;
    }

    public static void main(String[] args) {
        try {
            String path = new ClassPathResource("vicinitas_search_results_rihanna.xlsx").getFile().getAbsolutePath();
            FileInputStream file = new FileInputStream(new File(path));
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);
            Map<Integer, List<String>> data = new HashMap<>();
            int i = 0;
            for (Row row : sheet) {
                //data.put(i, new ArrayList<String>());
                String tweet = row.getCell(1).getRichStringCellValue().getString();
                //System.out.println(tweet);
                String text = tweet.trim()
                        // remove links
                        .replaceAll("http.*?[\\S]+", "")
                        // remove usernames
                        .replaceAll("@[\\S]+", "")
                        // replace hashtags by just words
                        .replaceAll("#", "")
                        // correct all multiple white spaces to a single white space
                        .replaceAll("[\\s]+", " ");

                int value = analyse(text);
                //System.out.println(value);
                List<String> temp = new ArrayList<String>();
                temp.add(text);
                /*
                    ["Very negative,", mapped["0"] || 0],
                    ["negative", mapped["1"] || 0],
                    ["neutral", mapped["2"] || 0],
                    ["positive", mapped["3"] || 0],
                    ["very positive", mapped["4"] || 0]
                */
                switch (value) {
                    case 0:
                        temp.add("Very negative");
                        break;
                    case 1:
                        temp.add("Negative");
                        break;
                    case 2:
                        temp.add("Neutral");
                        break;

                    case 3:
                        temp.add("Positive");
                        break;
                    case 4:
                        temp.add("Very positive");
                        break;
                }
                data.put(i, temp);
                i++;
                Cell headerCell = row.createCell(16);
                headerCell.setCellValue(value);
                //headerCell.setCellStyle(headerStyle);
            }
            writeData(data);

            FileOutputStream outputStream = new FileOutputStream(path);
            workbook.write(outputStream);
            workbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("DONE!");

    }

    private static void writeData(Map<Integer, List<String>> data) {
        try {
            String path = new ClassPathResource("test_rihanna.xlsx").getFile().getAbsolutePath();
            Workbook workbook = new XSSFWorkbook();

            Sheet sheet = workbook.createSheet("NLP");
            sheet.setColumnWidth(0, 6000);
            sheet.setColumnWidth(1, 4000);

            Row header = sheet.createRow(0);

            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            XSSFFont font = ((XSSFWorkbook) workbook).createFont();
            font.setFontName("Arial");
            font.setFontHeightInPoints((short) 16);
            font.setBold(true);
            headerStyle.setFont(font);

            Cell headerCell = header.createCell(0);
            headerCell.setCellValue("Tweet");
//            headerCell.setCellStyle(headerStyle);

            headerCell = header.createCell(1);
            headerCell.setCellValue("Sentiment");
//            headerCell.setCellStyle(headerStyle);

            int j = 1;
            for (Map.Entry<Integer, List<String>> entry : data.entrySet()) {
                writeCell(workbook, sheet, entry, j);
                j++;
            }

            FileOutputStream outputStream = new FileOutputStream(path);
            workbook.write(outputStream);
            outputStream.flush();
            workbook.close();
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void writeCell(Workbook workbook, Sheet sheet, Map.Entry<Integer, List<String>> entry,int rowIndex) {
        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);
        Row row = sheet.createRow(rowIndex);
        Cell cell = row.createCell(0);
        cell.setCellValue( entry.getValue().get(0));
        cell.setCellStyle(style);

        cell = row.createCell(1);
        cell.setCellValue(entry.getValue().get(1));
        cell.setCellStyle(style);
    }
}
