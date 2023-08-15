package com.heroku.java;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;

import javax.sql.DataSource;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.sql.Connection;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@SpringBootApplication
@Controller
public class GettingStartedApplication {
    private final DataSource dataSource;

    @Autowired
    public GettingStartedApplication(DataSource dataSource) {
        this.dataSource = dataSource;
    }

    @GetMapping("/")
    public String index() {
        return "index";
    }

    @GetMapping("/database")
    String database(Map<String, Object> model) {
        try (Connection connection = dataSource.getConnection()) {
            final var statement = connection.createStatement();
            statement.executeUpdate("CREATE TABLE IF NOT EXISTS ticks (tick timestamp)");
            statement.executeUpdate("INSERT INTO ticks VALUES (now())");

            final var resultSet = statement.executeQuery("SELECT tick FROM ticks");
            final var output = new ArrayList<>();
            while (resultSet.next()) {
                output.add("Read from DB: " + resultSet.getTimestamp("tick"));
            }

            model.put("records", output);
            return "database";

        } catch (Throwable t) {
            model.put("message", t.getMessage());
            return "error";
        }
    }
    @GetMapping("/download")
    public ResponseEntity<byte[]> downloadFile() throws IOException {
    Candidate candidate = new Candidate();
    candidate.setName("Nguyễn Huy Hoàng");

    String strTemplate = "Rikkeisoft_Skill_Sheet_Template_ja.xlsx";
    File currDir = new File(".");
    String path = currDir.getAbsolutePath();
    path = path.substring(0,path.length()-2);
        File filePath = new File(path + File.separator + PathHelper.RESOURCE_DIR  + File.separator + strTemplate);
        if (filePath.isFile()) {
//            byte[] fileOut = Files.readAllBytes(filePath.toPath());
            byte[] fileOut = ExcelHelper.excelExport(candidate, new FileInputStream(filePath), "EN");
            HttpHeaders headers = new HttpHeaders();
            String fileName = "abcd.xlsx";
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            headers.setContentDispositionFormData( fileName, fileName);
            return new ResponseEntity<>(fileOut, headers, HttpStatus.OK);
        }
        return new ResponseEntity<>(HttpStatus.NOT_FOUND);

    }
    public static void main(String[] args) {

        SpringApplication.run(GettingStartedApplication.class, args);
    }
}
