package org.difrango;

import com.beust.jcommander.Parameter;
import com.beust.jcommander.JCommander;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.Month;
import java.time.temporal.TemporalAdjusters;
import java.time.DayOfWeek;
import java.util.ArrayList;
import java.util.List;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.regex.Pattern;
import java.util.stream.StreamSupport;

public class WorkoutParser {
    private static final Pattern EXERCISE_PATTERN = Pattern.compile("^(\\d+)x(\\d+)(x(\\d+))?$");
    private static final Pattern NUMBER_PATTERN = Pattern.compile("\\d+");
    private static final Pattern DAY_PATTERN = Pattern.compile("Day \\d+");
    private static final Pattern WEEK_PATTERN = Pattern.compile("Week \\d+");

    @Parameter(names={"--folder", "-f"}, description = "The folder where the workout files are located")
    private String folder = "workouts";

    record MonthlyExercises(LocalDate date, List<WeeklyExercises> exercises) {
    }

    ;

    record WeeklyExercises(LocalDate date, List<DailyExercise> exercises) {
    }

    ;

    record DailyExercise(LocalDate date, String order, String name, int set, int reps, int weight) {
    }

    ;

    public static void main(String[] args) throws IOException {
        WorkoutParser main = new WorkoutParser();
        JCommander.newBuilder()
                .addObject(main)
                .build()
                .parse(args);
        main.run();
    }

    public void run() throws IOException {
        // Read the list of files
        var fileNames = Files.list(Paths.get(folder))
                .filter(Files::isRegularFile)
                .toList();
        // Create the output sheet
        try (var outputStream = new FileOutputStream("output.xlsx");
             var outputWorkbook = new XSSFWorkbook()) {
            var outputWorkbookSheet = outputWorkbook.createSheet("Monthly Exercises");
            var outputWorkbookSheetRow = outputWorkbookSheet.createRow(0);
            outputWorkbookSheetRow.createCell(0).setCellValue("Date");
            outputWorkbookSheetRow.createCell(1).setCellValue("Order");
            outputWorkbookSheetRow.createCell(2).setCellValue("Name");
            outputWorkbookSheetRow.createCell(3).setCellValue("Sets");
            outputWorkbookSheetRow.createCell(4).setCellValue("Reps");
            outputWorkbookSheetRow.createCell(5).setCellValue("Weight");

            // loop over the files
            var rowIndex = 1;
            for (var fileName : fileNames) {
                var excelFilePath = fileName.getFileName().toString();
                var month = Month.valueOf(StringUtils.upperCase(excelFilePath.split("-")[0]));
                var year = NumberUtils.toInt(excelFilePath.split("-")[1], 2024);
                var firstMonday =
                        LocalDate.of(year, month, 1).with(TemporalAdjusters.nextOrSame(DayOfWeek.MONDAY));

                var monthlyExercises = new MonthlyExercises(LocalDate.of(year, month, 1), new ArrayList<>());

                System.out.println(STR."\{excelFilePath} - The first Monday of \{month} \{year} is: \{firstMonday}");

                try (var inputWorkbook = new XSSFWorkbook(fileName.toFile())) {
                    // Get the first sheet (can also fetch by name)
                    StreamSupport.stream(Spliterators.spliteratorUnknownSize(inputWorkbook.sheetIterator(), Spliterator.ORDERED
                            ), false)
                            .takeWhile(sheet -> WEEK_PATTERN.matcher(sheet.getSheetName()).matches())
                            .forEach(sheet -> {
                                System.out.println(STR."Sheet Name is \{sheet.getSheetName()}");

                                var weeklyDate = firstMonday;
                                if (sheet.getSheetName().matches("Week [2-4]")) {
                                    var week = NUMBER_PATTERN.matcher(sheet.getSheetName()).results().findFirst().orElseThrow();
                                    weeklyDate = firstMonday.plusWeeks(Integer.parseInt(week.group()));
                                }

                                var weeklyExercises = new WeeklyExercises(weeklyDate, new ArrayList<>());
                                monthlyExercises.exercises.add(weeklyExercises);

                                sheet.rowIterator().forEachRemaining(row -> {
                                    var exerciseDate = weeklyExercises.date;
                                    var exercise = row.getCell(1);
                                    var set = row.getCell(5);

                                    if (exercise != null && set != null) {
                                        var exerciseValue = StringUtils.trim(exercise.getStringCellValue());
                                        var setCellValue = StringUtils.trim(set.getStringCellValue());

                                        var dayMatcher = DAY_PATTERN.matcher(exerciseValue);
                                        if (dayMatcher.find()) {
                                            var day = NUMBER_PATTERN.matcher(exerciseValue).results().findFirst().orElseThrow();
                                            exerciseDate = weeklyExercises.date.plusDays(Integer.parseInt(day.group()));
                                        }

                                        var setMatcher = EXERCISE_PATTERN.matcher(setCellValue);
                                        if (setMatcher.find()) {
                                            var exerciseSplit = exerciseValue.split(":");
                                            var dailyExercise = new DailyExercise(
                                                    exerciseDate, exerciseSplit[0], exerciseSplit[1],
                                                    NumberUtils.toInt(setMatcher.group(1), 0), NumberUtils.toInt(setMatcher.group(2), 0), NumberUtils.toInt(setMatcher.group(4), 0));
                                            weeklyExercises.exercises.add(dailyExercise);
                                        }
                                    }
                                });
                            });

                    System.out.println(STR."Monthly Exercises: \{monthlyExercises}");

                    for (var weeklyExercises : monthlyExercises.exercises) {
                        for (var dailyExercise : weeklyExercises.exercises) {
                            outputWorkbookSheetRow = outputWorkbookSheet.createRow(rowIndex++);
                            outputWorkbookSheetRow.createCell(0).setCellValue(dailyExercise.date.toString());
                            outputWorkbookSheetRow.createCell(1).setCellValue(dailyExercise.order);
                            outputWorkbookSheetRow.createCell(2).setCellValue(dailyExercise.name);
                            outputWorkbookSheetRow.createCell(3).setCellValue(dailyExercise.set);
                            outputWorkbookSheetRow.createCell(4).setCellValue(dailyExercise.reps);
                            outputWorkbookSheetRow.createCell(5).setCellValue(dailyExercise.weight);
                        }
                    }
                }
            }

            outputWorkbook.write(outputStream);
        } catch (IOException | InvalidFormatException e) {
            System.err.println(STR."Error reading the Excel file: \{e.getMessage()}");
        }
    }
}