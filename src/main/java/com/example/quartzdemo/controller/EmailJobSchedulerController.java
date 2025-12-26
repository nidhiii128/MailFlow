package com.example.quartzdemo.controller;

import com.example.quartzdemo.job.EmailJob;
import com.example.quartzdemo.payload.ScheduleEmailRequest;
import com.example.quartzdemo.payload.ScheduleEmailResponse;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.quartz.JobBuilder;
import org.quartz.JobDataMap;
import org.quartz.JobDetail;
import org.quartz.Scheduler;
import org.quartz.SchedulerException;
import org.quartz.Trigger;
import org.quartz.TriggerBuilder;
import org.quartz.SimpleScheduleBuilder;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import javax.validation.Valid;

import java.time.DateTimeException;
import java.time.LocalDateTime;
import java.time.ZonedDateTime;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.List;
import java.util.UUID;
import org.springframework.web.multipart.MultipartFile;
import java.time.ZoneId;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

@RestController
public class EmailJobSchedulerController {
    private static final Logger logger = LoggerFactory.getLogger(EmailJobSchedulerController.class);

    @Autowired
    private Scheduler scheduler;

    @PostMapping("/scheduleEmailsFromExcel")
    public ResponseEntity<List<ScheduleEmailResponse>> scheduleEmailsFromExcel(
            @RequestParam("file") MultipartFile file) {
        List<ScheduleEmailResponse> responses = new ArrayList<>();

        try {
            Workbook workbook = WorkbookFactory.create(file.getInputStream());
            Sheet sheet = workbook.getSheetAt(0);

            // Skip header row if it exists
            int startRow = sheet.getFirstRowNum() + 1;

            for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null)
                    continue;

                // Helper method to safely get cell value as string
                String email = getCellValueAsString(row.getCell(0));
                String dateTimeStr = getCellValueAsString(row.getCell(1));
                String salutation = getCellValueAsString(row.getCell(2));
                String company = getCellValueAsString(row.getCell(3));
                String name = getCellValueAsString(row.getCell(4));
                String designation = getCellValueAsString(row.getCell(5));
                String phone = getCellValueAsString(row.getCell(6));

                String subject = "Placement and Internship Invite 2025-26 | IIT Jodhpur | {company}";
                subject = subject.replace("{company}", company);
                String demoBody = "<!DOCTYPE html><html lang=\\\"en\\\"><head><meta charset=\\\"UTF-8\\\"><meta name=\\\"viewport\\\" content=\\\"width=device-width, initial-scale=1.0\\\"><title>IIT Jodhpur Placement Letter</title><style>body{font-family:Arial,sans-serif;line-height:1.6;max-width:800px;margin:20px auto;padding:0 20px}table{width:100%;border-collapse:collapse;margin:20px 0}th,td{border:1px solid #ddd;padding:8px;text-align:left}th{background-color:#b8cce4}.header{margin-bottom:20px}.section-title{font-weight:bold;margin-top:20px}.contact-info{font-style:italic;margin-top:20px}ul{padding-left:20px}</style></head><body><div class=\\\"header\\\"><p>Dear {salutation},</p><p>Greetings from IIT Jodhpur!</p></div><p>On behalf of the Placement Cell at <strong>IIT Jodhpur</strong>, I, Arman Gupta, Internship Coordinator, take this opportunity to invite <strong>{company}</strong> to participate in our campus placement and internship season for the 2025 and 2026 batches, respectively.</p><p>Since its inception in 2008, IIT Jodhpur has achieved several milestones and has always strived to push its limits in all spheres. The institute has a large pool of talented students pursuing their interests through a wide range of academic programs. Notably, IIT Jodhpur secured the <strong>29th rank</strong> in NIRF 2024.</p><p>IIT Jodhpur stands out with its <strong>Industry 4.0 curriculum</strong>, interdisciplinary projects, and mandatory courses in Machine Learning and Data Structures for all branches. Our students are actively engaged in various tech and non-tech clubs ensuring they are well-rounded and industry-ready.</p><div class=\\\"section-title\\\">Why Collaborate with IIT Jodhpur?</div><ul><li><strong>Qualified Talent Pool:</strong> Our students undergo rigorous training and excel both academically and practically.</li><li><strong>Diverse Skill Sets:</strong> Programs offered include B.Tech, BS, M.Tech, M.Sc, Ph.D., and dual degrees across various departments.</li><li><strong>Innovative Learning:</strong> Our curriculum is updated with the latest industry trends and technologies, focusing on emerging domains like Artificial Intelligence, IoT, and Computational Sciences.</li><li><strong>Interdisciplinary Projects & Research:</strong> Students engage in projects that integrate multiple disciplines, preparing them for complex industry challenges.</li><li><strong>Active Clubs:</strong> Tech and non-tech clubs, such as Product, DevLup Labs, RAID, Robotics Society, and E-Cell, contribute to the holistic development of our students.</li></ul><div class=\\\"section-title\\\">PLACEMENTS</div><table><tr><th>Programs Offered</th><th>Available Batch Strength</th></tr><tr><td>B.Tech</td><td>440</td></tr><tr><td>M.Tech</td><td>210</td></tr><tr><td>M.Sc</td><td>80</td></tr><tr><td>Tech MBA</td><td>80</td></tr></table><div class=\\\"section-title\\\">INTERNSHIPS</div><table><tr><th>Programs Offered</th><th>Available Batch Strength</th></tr><tr><td>B.Tech</td><td>500</td></tr><tr><td>Tech MBA</td><td>120</td></tr></table><p>For more details, please find the <a href=\\\"#\\\">brochure</a> attached. We invite you to consider our students for both technical and non-technical roles. Kindly fill out and return the attached <a href=\\\"#\\\">Job (JAF)</a> / <a href=\\\"#\\\">Internship (IAF)</a> Announcement Form to expedite the process.</p><p>We look forward to a long-term relationship with your organization. For any queries, feel free to contact me or our team.</p><div class=\\\"contact-info\\\"><p>Warm Regards,<br>Arman Gupta<br>Internship Coordinator<br>Career Development Cell | IIT Jodhpur<br>Contact : +91 8273072067</p><p>Alternate Contact -<br>Puneet Garg<br>Training & Placement Officer<br>Career Development Cell | IIT Jodhpur<br>Contact : +91 9815964823 , 0291-2801155</p></div></body></html>";

                String body = demoBody.replace("{salutation}", salutation);
                body = body.replace("{company}", company);
                body = body.replace("{name}", name);
                body = body.replace("{designation}", designation);
                body = body.replace("{phone}", phone);

                String timeZoneStr = "Asia/Kolkata";

                // Skip if any required field is empty
                if (email.isEmpty() || subject.isEmpty() || body.isEmpty() || dateTimeStr.isEmpty()
                        || timeZoneStr.isEmpty()) {
                    logger.warn("Skipping row {} due to missing required fields", i + 1);
                    continue;
                }

                // Check if the datetime is in the future
                LocalDateTime dateTime;
                try {
                    dateTime = LocalDateTime.parse(dateTimeStr);
                    if (dateTime.isBefore(LocalDateTime.now())) {
                        logger.warn("Skipping row {} due to past dateTime: {}", i + 1, dateTimeStr);
                        continue;
                    }
                } catch (DateTimeParseException e) {
                    logger.error("Invalid date time format in row {}: {}", i + 1, dateTimeStr);
                    continue;
                }

                ZoneId timeZone;
                try {
                    timeZone = ZoneId.of(timeZoneStr);
                } catch (DateTimeException e) {
                    logger.error("Invalid time zone in row {}: {}", i + 1, timeZoneStr);
                    continue;
                }

                ScheduleEmailRequest scheduleEmailRequest = new ScheduleEmailRequest();
                scheduleEmailRequest.setEmail(email);
                scheduleEmailRequest.setDateTime(dateTime);
                scheduleEmailRequest.setSalutation(salutation);
                scheduleEmailRequest.setCompany(company);
                scheduleEmailRequest.setName(name);
                scheduleEmailRequest.setDesignation(designation);
                scheduleEmailRequest.setPhone(phone);

                ScheduleEmailResponse scheduleEmailResponse = demo(scheduleEmailRequest);
                responses.add(scheduleEmailResponse);
                System.out.println(scheduleEmailResponse);
            }

            workbook.close();
            return ResponseEntity.ok(responses);

        } catch (Exception ex) {
            logger.error("Error processing Excel file", ex);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body(Collections.singletonList(
                            new ScheduleEmailResponse(false, "Error processing Excel file: " + ex.getMessage())));
        }
    }



    // Helper method to handle different cell types
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    // Handle date cells
                    return cell.getLocalDateTimeCellValue().toString();
                }
                // Convert numeric value to string with proper formatting
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return String.valueOf(cell.getNumericCellValue());
                } catch (IllegalStateException e) {
                    return cell.getStringCellValue();
                }
            case BLANK:
                return "";
            default:
                return "";
        }
    }

    private ScheduleEmailResponse demo(ScheduleEmailRequest scheduleEmailRequest) {
        try {
            ZonedDateTime dateTime = ZonedDateTime.of(scheduleEmailRequest.getDateTime(),
                    ZoneId.of("Asia/Kolkata"));
            if (dateTime.isBefore(ZonedDateTime.now())) {
                ScheduleEmailResponse scheduleEmailResponse = new ScheduleEmailResponse(false,
                        "dateTime must be after current time");
                return scheduleEmailResponse;
            }

            JobDetail jobDetail = buildJobDetail(scheduleEmailRequest);
            Trigger trigger = buildJobTrigger(jobDetail, dateTime);
            scheduler.scheduleJob(jobDetail, trigger);

            ScheduleEmailResponse scheduleEmailResponse = new ScheduleEmailResponse(true,
                    jobDetail.getKey().getName(), jobDetail.getKey().getGroup(), "Email Scheduled Successfully!");
            return scheduleEmailResponse;
        } catch (SchedulerException ex) {
            logger.error("Error scheduling email", ex);

            ScheduleEmailResponse scheduleEmailResponse = new ScheduleEmailResponse(false,
                    "Error scheduling email. Please try later!");
            return scheduleEmailResponse;
        }
    }

    @PostMapping("/scheduleEmail")
    public ResponseEntity<ScheduleEmailResponse> scheduleEmail(
            @Valid @RequestBody ScheduleEmailRequest scheduleEmailRequest) {
        try {
            ZonedDateTime dateTime = ZonedDateTime.of(scheduleEmailRequest.getDateTime(),
                    ZoneId.of("Asia/Kolkata"));
            System.out.println(dateTime);
            if (dateTime.isBefore(ZonedDateTime.now())) {
                ScheduleEmailResponse scheduleEmailResponse = new ScheduleEmailResponse(false,
                        "dateTime must be after current time");
                return ResponseEntity.badRequest().body(scheduleEmailResponse);
            }

            JobDetail jobDetail = buildJobDetail(scheduleEmailRequest);
            Trigger trigger = buildJobTrigger(jobDetail, dateTime);
            scheduler.scheduleJob(jobDetail, trigger);

            ScheduleEmailResponse scheduleEmailResponse = new ScheduleEmailResponse(true,
                    jobDetail.getKey().getName(), jobDetail.getKey().getGroup(), "Email Scheduled Successfully!");
            return ResponseEntity.ok(scheduleEmailResponse);
        } catch (SchedulerException ex) {
            logger.error("Error scheduling email", ex);

            ScheduleEmailResponse scheduleEmailResponse = new ScheduleEmailResponse(false,
                    "Error scheduling email. Please try later!");
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(scheduleEmailResponse);
        }
    }

    private JobDetail buildJobDetail(ScheduleEmailRequest scheduleEmailRequest) {

        String company = scheduleEmailRequest.getCompany();
        String salutation = scheduleEmailRequest.getSalutation();
        String name = scheduleEmailRequest.getName();
        String designation = scheduleEmailRequest.getDesignation();
        String phone = scheduleEmailRequest.getPhone();

        String subject = "Placement and Internship Invite 2025-26 | IIT Jodhpur | {company}";
                subject = subject.replace("{company}", company);
                String demoBody = "<!DOCTYPE html><html lang=\\\"en\\\"><head><meta charset=\\\"UTF-8\\\"><meta name=\\\"viewport\\\" content=\\\"width=device-width, initial-scale=1.0\\\"><title>IIT Jodhpur Placement Letter</title><style>body{font-family:Arial,sans-serif;line-height:1.6;max-width:800px;margin:20px auto;padding:0 20px}table{width:100%;border-collapse:collapse;margin:20px 0}th,td{border:1px solid #ddd;padding:8px;text-align:left}th{background-color:#b8cce4}.header{margin-bottom:20px}.section-title{font-weight:bold;margin-top:20px}.contact-info{font-style:italic;margin-top:20px}ul{padding-left:20px}</style></head><body><div class=\\\"header\\\"><p>Dear {salutation},</p><p>Greetings from IIT Jodhpur!</p></div><p>On behalf of the Placement Cell at <strong>IIT Jodhpur</strong>, I, {name}, {designation}, take this opportunity to invite <strong>{company}</strong> to participate in our campus placement and internship season for the 2025 and 2026 batches, respectively.</p><p>Since its inception in 2008, IIT Jodhpur has achieved several milestones and has always strived to push its limits in all spheres. The institute has a large pool of talented students pursuing their interests through a wide range of academic programs. Notably, IIT Jodhpur secured the <strong>29th rank</strong> in NIRF 2024.</p><p>IIT Jodhpur stands out with its <strong>Industry 4.0 curriculum</strong>, interdisciplinary projects, and mandatory courses in Machine Learning and Data Structures for all branches. Our students are actively engaged in various tech and non-tech clubs ensuring they are well-rounded and industry-ready.</p><div class=\\\"section-title\\\">Why Collaborate with IIT Jodhpur?</div><ul><li><strong>Qualified Talent Pool:</strong> Our students undergo rigorous training and excel both academically and practically.</li><li><strong>Diverse Skill Sets:</strong> Programs offered include B.Tech, BS, M.Tech, M.Sc, Ph.D., and dual degrees across various departments.</li><li><strong>Innovative Learning:</strong> Our curriculum is updated with the latest industry trends and technologies, focusing on emerging domains like Artificial Intelligence, IoT, and Computational Sciences.</li><li><strong>Interdisciplinary Projects & Research:</strong> Students engage in projects that integrate multiple disciplines, preparing them for complex industry challenges.</li><li><strong>Active Clubs:</strong> Tech and non-tech clubs, such as Product, DevLup Labs, RAID, Robotics Society, and E-Cell, contribute to the holistic development of our students.</li></ul><div class=\\\"section-title\\\">PLACEMENTS</div><table><tr><th>Programs Offered</th><th>Available Batch Strength</th></tr><tr><td>B.Tech</td><td>440</td></tr><tr><td>M.Tech</td><td>210</td></tr><tr><td>M.Sc</td><td>80</td></tr><tr><td>Tech MBA</td><td>80</td></tr></table><div class=\\\"section-title\\\">INTERNSHIPS</div><table><tr><th>Programs Offered</th><th>Available Batch Strength</th></tr><tr><td>B.Tech</td><td>500</td></tr><tr><td>Tech MBA</td><td>120</td></tr></table><p>For more details, please find the <a href=\\\"#\\\">brochure</a> attached. We invite you to consider our students for both technical and non-technical roles. Kindly fill out and return the attached <a href=\\\"#\\\">Job (JAF)</a> / <a href=\\\"#\\\">Internship (IAF)</a> Announcement Form to expedite the process.</p><p>We look forward to a long-term relationship with your organization. For any queries, feel free to contact me or our team.</p><div class=\\\"contact-info\\\"><p>Warm Regards,<br>{name}<br>{designation}<br>Career Development Cell | IIT Jodhpur<br>Contact : +91 {phone}</p><p>Alternate Contact -<br>Puneet Garg<br>Training & Placement Officer<br>Career Development Cell | IIT Jodhpur<br>Contact : +91 9815964823 , 0291-2801155</p></div></body></html>";

                String body = demoBody.replace("{salutation}", salutation);
                body = body.replace("{company}", company);
                body = body.replace("{name}", name);
                body = body.replace("{designation}", designation);
                body = body.replace("{phone}", phone);
                String timeZoneStr = "Asia/Kolkata";

        JobDataMap jobDataMap = new JobDataMap();

        jobDataMap.put("email", scheduleEmailRequest.getEmail());
        jobDataMap.put("subject", subject);
        jobDataMap.put("body", body);

        return JobBuilder.newJob(EmailJob.class)
                .withIdentity(UUID.randomUUID().toString(), "email-jobs")
                .withDescription("Send Email Job")
                .usingJobData(jobDataMap)
                .storeDurably()
                .build();
    }

    private Trigger buildJobTrigger(JobDetail jobDetail, ZonedDateTime startAt) {
        return TriggerBuilder.newTrigger()
                .forJob(jobDetail)
                .withIdentity(jobDetail.getKey().getName(), "email-triggers")
                .withDescription("Send Email Trigger")
                .startAt(Date.from(startAt.toInstant()))
                .withSchedule(SimpleScheduleBuilder.simpleSchedule().withMisfireHandlingInstructionFireNow())
                .build();
    }
}
