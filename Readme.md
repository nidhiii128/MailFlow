# 📧 Automated Email Scheduler

A robust full-stack application built with **Spring Boot** and **Quartz Scheduler** that allows for precise scheduling of personalized emails. This project supports both single entries and bulk processing via Excel, persisting all jobs in a local **MySQL** database.



---

## 🚀 Key Features

* **Quartz Persistence:** Jobs are saved in MySQL tables, ensuring they are not lost if the server restarts.
* **Single Email Scheduling:** A dedicated dashboard to schedule emails with custom templates (Company Name, Designation, etc.).
* **Bulk Processing:** Upload `.xlsx` or `.xls` files to schedule hundreds of emails simultaneously.
* **Gmail SMTP Integration:** Secured via Google App Passwords for reliable delivery.
* **Modern UI:** Responsive dashboard served via Thymeleaf.

---

## 🛠️ Tech Stack

* **Backend:** Java 17, Spring Boot 2.5.5
* **Scheduler:** Quartz Scheduler 2.3.2
* **Database:** MySQL 8.0
* **Frontend:** Thymeleaf, HTML5, CSS3 (Modern Dashboard UI)
* **Build Tool:** Maven

---

## ⚙️ Local Setup Instructions

### 1. Database Configuration
1. Open your MySQL Command Line.
2. Create a new database:
   ```sql
   CREATE DATABASE quartz_demo;

Quartz will automatically generate its required QRTZ_ tables upon the first successful run.

2. Configure Environment
Update src/main/resources/application.properties with your local credentials:

Properties

# MySQL Connection
spring.datasource.url=jdbc:mysql://localhost:3306/quartz_demo
spring.datasource.username=YOUR_USERNAME
spring.datasource.password=YOUR_PASSWORD

# Gmail SMTP (Google App Password)
spring.mail.username=your-email@gmail.com
spring.mail.password=your-16-character-app-code
3. Running the Application
Open your terminal in the project root and run:

PowerShell

.\mvnw spring-boot:run
The application will be accessible at: http://localhost:8080

📂 Project Structure
src/main/java: Contains the Quartz Job logic and Controller endpoints.

src/main/resources/templates: Contains the index.html UI.

src/main/resources/static: Contains images and CSS assets.

pom.xml: Managed dependencies including Spring Mail and Quartz.

<img width="1896" height="919" alt="image" src="https://github.com/user-attachments/assets/c26f495c-ce16-41f3-8255-42aee23dddec" />
