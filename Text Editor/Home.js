/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
    See full license at the bottom of this file. */

(function () {
	"use strict";

	var courses;
	var catalog;

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function () {
			app.initialize();

			$("#tabs").tabs();

			$('#generate-planner').button();
			$('#generate-planner').click(createCollegeCreditTracker);

			$('#add-course').button();
			$('#add-course').click(addCourseToTable);

			$('#semester').selectmenu();


			// Typically this data would come from a web service
			// Sample courses for initial report
			courses = [
							 {
								 "course_title": "Anthropology",
								 "course_no": "GEN 108",
								 "degree_req": "General Study",
								 "credits": "5",
								 "completed": "Yes",
								 "semester": "Semester 1"
							 },
							{
								"course_title": "Applied Music",
								"course_no": "MUS 215",
								"degree_req": "Academic Major",
								"credits": "4",
								"completed": "No",
								"semester": "Semester 3"
							},
							{
								"course_title": "Art History I",
								"course_no": "ART 101",
								"degree_req": "General Study",
								"credits": "4",
								"completed": "Yes",
								"semester": "Semester 1"
							},
							{
								"course_title": "Art History II",
								"course_no": "ART 201",
								"degree_req": "General Study",
								"credits": "5",
								"completed": "Yes",
								"semester": "Semester 2"
							},
							{
								"course_title": "Aural Skills I",
								"course_no": "MUS 113",
								"degree_req": "Academic Major",
								"credits": "5",
								"completed": "Yes",
								"semester": "Semester 1"
							},
							{
								"course_title": "Aural Skills II",
								"course_no": "MUS 213",
								"degree_req": "Academic Major",
								"credits": "5",
								"completed": "Yes",
								"semester": "Semester 2"
							},
							{
								"course_title": "Aural Skills III",
								"course_no": "MUS 313",
								"degree_req": "Academic Major",
								"credits": "5",
								"completed": "No",
								"semester": "Semester 3"
							},
							{
								"course_title": "Aural Skills IV",
								"course_no": "MUS 413",
								"degree_req": "Academic Major",
								"credits": "5",
								"completed": "No",
								"semester": "Semester 4"
							},
							{
								"course_title": "Conducting I",
								"course_no": "MUS 114",
								"degree_req": "Academic Major",
								"credits": "2",
								"completed": "Yes",
								"semester": "Semester 1"
							},
							{
								"course_title": "English Writing",
								"course_no": "Eng 101",
								"degree_req": "Generic Study",
								"credits": "3",
								"completed": "Yes",
								"semester": "Semester 1"
							},
							{
								"course_title": "English Writing",
								"course_no": "Eng 201",
								"degree_req": "Generic Study",
								"credits": "3",
								"completed": "Yes",
								"semester": "Semester 2"
							},
							{
								"course_title": "Form and Analysis",
								"course_no": "MUS 214",
								"degree_req": "Academic Major",
								"credits": "2",
								"completed": "Yes",
								"semester": "Semester 2"
							},
							{
								"course_title": "Intro to Anthropology",
								"course_no": "GEN 208",
								"degree_req": "General Study",
								"credits": "3",
								"completed": "Yes",
								"semester": "Semester 2"
							},
							{
								"course_title": "Mathematics 101",
								"course_no": "MAT 101",
								"degree_req": "General Study",
								"credits": "5",
								"completed": "Yes",
								"semester": "Semester 1"
							},
							{
								"course_title": "Music History in Western Civilization",
								"course_no": "MUS 101",
								"degree_req": "Academic Major",
								"credits": "2",
								"completed": "Yes",
								"semester": "Semester 1"
							},
							{
								"course_title": "Music History in Western Civilization",
								"course_no": "MUS 201",
								"degree_req": "Academic Major",
								"credits": "2",
								"completed": "Yes",
								"semester": "Semester 1"
							},
							{
								"course_title": "Music Theory I",
								"course_no": "MUS 110",
								"degree_req": "Academic Major",
								"credits": "2",
								"completed": "Yes",
								"semester": "Semester 2"
							},
							{
								"course_title": "Music Theory II",
								"course_no": "MUS 210",
								"degree_req": "Academic Major",
								"credits": "2",
								"completed": "Yes",
								"semester": "Semester 3"
							},
							{
								"course_title": "Music Theory III",
								"course_no": "MUS 310",
								"degree_req": "Academic Major",
								"credits": "2",
								"completed": "No",
								"semester": "Semester 4"
							},
							{
								"course_title": "Music Theory IV",
								"course_no": "MUS 410",
								"degree_req": "Academic Major",
								"credits": "2",
								"completed": "No",
								"semester": "Semester 5"
							},
							{
								"course_title": "Piano Class",
								"course_no": "MUS 109",
								"degree_req": "Academic Major",
								"credits": "2",
								"completed": "Yes",
								"semester": "Semester 1"
							},
							{
								"course_title": "Social Sciences 101",
								"course_no": "SOC 101",
								"degree_req": "General Study",
								"credits": "3",
								"completed": "Yes",
								"semester": "Semester 1"
							},
							{
								"course_title": "Social Studies 101",
								"course_no": "SOC 201",
								"degree_req": "General Study",
								"credits": "3",
								"completed": "Yes",
								"semester": "Semester 1"
							},
							{
								"course_title": "World of Jazz",
								"course_no": "MUS 105",
								"degree_req": "Elective Course",
								"credits": "4",
								"completed": "Yes",
								"semester": "Semester 2"
							},
							{
								"course_title": "World of Music I",
								"course_no": "MUS 112",
								"degree_req": "Academic Major",
								"credits": "6",
								"completed": "Yes",
								"semester": "Semester 3"
							},
							{
								"course_title": "World of Music II",
								"course_no": "MUS 212",
								"degree_req": "Academic Major",
								"credits": "6",
								"completed": "Yes",
								"semester": "Semester 4"
							},
							{
								"course_title": "World of Music III",
								"course_no": "MUS 312",
								"degree_req": "Academic Major",
								"credits": "6",
								"completed": "Yes",
								"semester": "Semester 5"
							},
							{
								"course_title": "World of Music IV",
								"course_no": "MUS 412",
								"degree_req": "Academic Major",
								"credits": "6",
								"completed": "Yes",
								"semester": "Semester 6"
							},
							{
								"course_title": "World of Music V",
								"course_no": "MUS 512",
								"degree_req": "Academic Major",
								"credits": "6",
								"completed": "Yes",
								"semester": "Semester 7"
							},
							{
								"course_title": "World of Music VI",
								"course_no": "MUS 512",
								"degree_req": "Academic Major",
								"credits": "6",
								"completed": "Yes",
								"semester": "Semester 8"
							}
			];

			// Typically this data would come from a web service
			//  Sample course catalog
			catalog = [
							{
								"course_title": "Survey of World Music",
								"course_no": "MUS 302",
								"degree_req": "Academic Major",
								"credits": "5"
							},
							{
								"course_title": "Music and the Mind",
								"course_no": "MUS 219",
								"degree_req": "Academic Major",
								"credits": "4"
							},
							{
								"course_title": "Factors of Musical Learning",
								"course_no": "MUS 501",
								"degree_req": "Academic Major",
								"credits": "3"
							},
							{
								"course_title": "Art History III",
								"course_no": "ART 301",
								"degree_req": "General Study",
								"credits": "5",
							},
							{
								"course_title": "Music Business Internship",
								"course_no": "ART 301",
								"degree_req": "Academic Major",
								"credits": "10",
							},
							{
								"course_title": "Hist and Philosophy of Music Ed",
								"course_no": "GEN 101",
								"degree_req": "Generic Study",
								"credits": "3",
							},
							{
								"course_title": "Music History and Literature",
								"course_no": "MUS 208",
								"degree_req": "Academic Major",
								"credits": "5",
							},
							{
								"course_title": "Thesis in Music",
								"course_no": "MUS 302",
								"degree_req": "Academic Major",
								"credits": "8",
							},
							{
								"course_title": "Sports in Secondary Schools",
								"course_no": "ELE 176",
								"degree_req": "Elective Course",
								"credits": "3",
							},
							{
								"course_title": "Ed for Classroom Teachers",
								"course_no": "ELE 198",
								"degree_req": "Elective Course",
								"credits": "6",
							}
			];

			// Populate the drop down list with the course catalog info
			for (var i in catalog) {
				// Add each course to the drop down list
				$("#course-name").append("<option value=" + i + ">" + catalog[i].course_title + "</option>");
			}

			// Pouplate the other fields in the Add course tab when a course is selected
			$('#course-name').selectmenu({
				change: function (event, ui) {
					$("#course-id").val(catalog[$("#course-name").val()].course_no);
					$("#credit-type").val(catalog[$("#course-name").val()].degree_req);
					$("#credits").val(catalog[$("#course-name").val()].credits);
				}
			});
		});
	};

	// Create the tracker in Excel
	function createCollegeCreditTracker() {


		// Run a batch operation against the Excel object model
		Excel.run(function (ctx) {

			// Create a proxy object for the active worksheet
			var dashboardSheet = ctx.workbook.worksheets.getActiveWorksheet();

			// Queue a command to clear the contents before inserting data
			dashboardSheet.getUsedRange().clear();

			// Queue a command to rename the sheet
			dashboardSheet.name = "Dashboard";

			// Queue commands to set the title and format it
			var title = "College Credit Planner";
			var range = dashboardSheet.getRange("A1");
			range.values = title;
			range.format.font.name = "Century";
			range.format.font.size = 26;
			range.format.font.color = "white";

			// Queue a command to fill the title row
			dashboardSheet.getRange("A1:K1").format.fill.color = "1E8FEB";

			// Queue commands to add the degree next to the title
			var degreeName = "Bachelor of Arts in Music History";
			var range = dashboardSheet.getRange("C1");
			range.values = degreeName;
			range.format.font.name = "Century";
			range.format.font.size = 14;
			range.format.font.color = "white";

			// Queue commands to add the College Courses table with sample data
			// Queue commands to set the title for the table
			var title = "College Courses";
			var range = dashboardSheet.getRange("A23");
			range.values = title;
			range.format.font.name = "Century";
			range.format.font.size = 26;
			range.format.font.color = "white";

			dashboardSheet.getRange("A23:K23").format.fill.color = "1E8FEB";

			// Queue a command to add a new table
			var table = ctx.workbook.tables.add('Dashboard!A24:F24', true);
			table.name = "coursesTable";

			// Queue a command to get the newly added table
			table.getHeaderRowRange().values = [["COURSE TITLE", "COURSE#", "DEGREE REQUIREMENT", "CREDITS", "COMPLETED?", "SEMESTER"]];

			// Create a proxy object for the table rows
			var tableRows = table.rows;

			for (var i in courses) {
				// Queue commands to add some sample rows to the course table
				tableRows.add(null, [[courses[i].course_title, courses[i].course_no, courses[i].degree_req, courses[i].credits, courses[i].completed, courses[i].semester]]);
			}

			// Queue commands to format the table
			var range = dashboardSheet.getRange("A24:I24");
			range.format.font.size = 11;
			range.format.font.name = "Franklin Gothic Medium";
			range.format.font.color = "white";
			range.format.fill.color = "2A4C69";
			dashboardSheet.getRange("A25:I125").format.font.name = "Bookman Old Style";
			dashboardSheet.getRange("A25:I125").format.font.size = 10;
			dashboardSheet.getRange("A24:I125").format.borders.getItem('EdgeBottom').style = 'Continuous';
			dashboardSheet.getRange("A24:I125").format.borders.getItem('EdgeTop').style = 'Continuous';


			// Queue commands to add the CREDITS section in the sheet
			var creditreqvalues = [["CREDIT REQUIREMENTS", "TOTAL", "EARNED", "NEEDED"],
						   ["Academic Major", 90, '=IFERROR(SUMIFS(D22:D122,C22:C122,C4,E22:E122,"=Yes"),"")', "=D4-E4"],
						   ["Academic Minor", 0, '=IFERROR(SUMIFS(D22:D122,C22:C122,C5,E22:E122,"=Yes"),"")', "=D5-E5"],
						   ["Elective Course", 8, '=IFERROR(SUMIFS(D22:D122,C22:C122,C6,E22:E122,"=Yes"),"")', "=D6-E6"],
						   ["General Study", 66, '=IFERROR(SUMIFS(D22:D122,C22:C122,C7,E22:E122,"=Yes"),"")', "=D7-E7"],
						   ["Totals", "=SUM(D4:D7)", "=SUM(E4:E7)", "=SUM(F4:F7)"]];
			dashboardSheet.getRange("C3:F8").values = creditreqvalues;

			// Queue commands to format the CREDITS section
			dashboardSheet.getRange("C3:F3").format.font.size = 11;
			dashboardSheet.getRange("C3:F3").format.font.name = "Franklin Gothic Medium";
			dashboardSheet.getRange("C4:F8").format.font.name = "Bookman Old Style";
			dashboardSheet.getRange("C4:F7").format.font.size = 10;
			dashboardSheet.getRange("C3:F8").format.borders.getItem("InsideHorizontal").style = "Continuous";
			dashboardSheet.getRange("C3:F8").format.borders.getItem('EdgeBottom').style = 'Continuous';
			dashboardSheet.getRange("C3:F8").format.borders.getItem('EdgeTop').style = 'Continuous';
			dashboardSheet.getRange("C8:F8").format.font.size = 11;
			dashboardSheet.getRange("C8:F8").format.font.name = "Franklin Gothic Medium";


			// Queue commands to add the SEMESTER CREDITS and CLASSES section
			var semestersummarydata = [["SEMESTER", "CREDITS", "CLASSES"],
						   ["Semester 1", '=IFERROR(SUMIFS(D22:D122,F22:F122,C12),"")', '=IFERROR(COUNTIFS(F22:F122,C12),"")'],
						   ["Semester 2", '=IFERROR(SUMIFS(D22:D122,F22:F122,C13),"")', '=IFERROR(COUNTIFS(F22:F122,C13),"")'],
						   ["Semester 3", '=IFERROR(SUMIFS(D22:D122,F22:F122,C14),"")', '=IFERROR(COUNTIFS(F22:F122,C14),"")'],
						   ["Semester 4", '=IFERROR(SUMIFS(D22:D122,F22:F122,C15),"")', '=IFERROR(COUNTIFS(F22:F122,C15),"")'],
						   ["Semester 5", '=IFERROR(SUMIFS(D22:D122,F22:F122,C16),"")', '=IFERROR(COUNTIFS(F22:F122,C16),"")'],
						   ["Semester 6", '=IFERROR(SUMIFS(D22:D122,F22:F122,C17),"")', '=IFERROR(COUNTIFS(F22:F122,C17),"")'],
						   ["Semester 7", '=IFERROR(SUMIFS(D22:D122,F22:F122,C18),"")', '=IFERROR(COUNTIFS(F22:F122,C18),"")'],
						   ["Semester 8", '=IFERROR(SUMIFS(D22:D122,F22:F122,C19),"")', '=IFERROR(COUNTIFS(F22:F122,C19),"")'],
						   ["Total", "=sum(D12:D19)", "=sum(E12:E19)"]];
			dashboardSheet.getRange("C11:E20").values = semestersummarydata;

			// Queue commands to format the SEMESTER section
			dashboardSheet.getRange("C11:E11").format.font.size = 11;
			dashboardSheet.getRange("C11:E11").format.font.name = "Franklin Gothic Medium";
			dashboardSheet.getRange("C12:E19").format.font.name = "Bookman Old Style";
			dashboardSheet.getRange("C12:E19").format.font.size = 10;
			dashboardSheet.getRange("C11:E20").format.borders.getItem("InsideHorizontal").style = "Continuous";
			dashboardSheet.getRange("C11:E20").format.borders.getItem('EdgeBottom').style = 'Continuous';
			dashboardSheet.getRange("C11:E20").format.borders.getItem('EdgeTop').style = 'Continuous';
			dashboardSheet.getRange("C20:E20").format.font.size = 11;
			dashboardSheet.getRange("C20:E20").format.font.name = "Franklin Gothic Medium";
			dashboardSheet.getRange("C20:E20").format.font.color = "white";
			dashboardSheet.getRange("C20:E20").format.fill.color = "2A4C69";


			// Queue commands to create the chart based on the semester data
			// Queue commands to set the title for the chart
			var charttitle = "SEMESTER SUMMARY";
			dashboardSheet.getRange("A3:A3").values = charttitle;
			dashboardSheet.getRange("A3:A3").format.font.name = "Franklin Gothic Medium";
			dashboardSheet.getRange("A3:A3").format.font.size = 11;
			var semestersummarytitle = "Semester Summary Data";
			dashboardSheet.getRange("C10:C10").values = semestersummarytitle;
			dashboardSheet.getRange("C10:C10").format.font.name = "Century";
			dashboardSheet.getRange("C10:E10").format.font.size = 11;
			dashboardSheet.getRange("C10:E10").format.fill.color = "1E8FEB";
			dashboardSheet.getRange("C10:E10").format.font.color = "white";

			// Queue commands to create a chart and format it
			var chartDataRange = dashboardSheet.getRange("C11:E19");
			var semestersummarychart = dashboardSheet.charts.add("BarClustered", chartDataRange, "auto");
			semestersummarychart.setPosition("A4", "A19");
			semestersummarychart.legend.position = "right";
			semestersummarychart.title.visible = false;
			semestersummarychart.dataLabels.showValue = true;
			semestersummarychart.series.getItemAt(0).format.fill.setSolidColor("green");

			// Run the queued-up commands, and return a promise to indicate task completion
			return ctx.sync();
			// Queue commands to format the table
			var range = dashboardSheet.getRange("A24:I24");
			range.format.font.size = 11;
			range.format.font.name = "Franklin Gothic Medium";
			range.format.font.color = "white";
			range.format.fill.color = "2A4C69";
			dashboardSheet.getRange("A25:I125").format.font.name = "Bookman Old Style";
			dashboardSheet.getRange("A25:I125").format.font.size = 10;
			dashboardSheet.getRange("A24:I125").format.borders.getItem('EdgeBottom').style = 'Continuous';
			dashboardSheet.getRange("A24:I125").format.borders.getItem('EdgeTop').style = 'Continuous';


			// Queue commands to add the CREDITS section in the sheet
			var creditreqvalues = [["CREDIT REQUIREMENTS", "TOTAL", "EARNED", "NEEDED"],
						   ["Academic Major", 90, '=IFERROR(SUMIFS(D22:D122,C22:C122,C4,E22:E122,"=Yes"),"")', "=D4-E4"],
						   ["Academic Minor", 0, '=IFERROR(SUMIFS(D22:D122,C22:C122,C5,E22:E122,"=Yes"),"")', "=D5-E5"],
						   ["Elective Course", 4, '=IFERROR(SUMIFS(D22:D122,C22:C122,C6,E22:E122,"=Yes"),"")', "=D6-E6"],
						   ["General Study", 66, '=IFERROR(SUMIFS(D22:D122,C22:C122,C7,E22:E122,"=Yes"),"")', "=D7-E7"],
						   ["Totals", "=SUM(D4:D7)", "=SUM(E4:E7)", "=SUM(F4:F7)"]];
			dashboardSheet.getRange("C3:F8").values = creditreqvalues;

			// Queue commands to format the CREDITS section
			dashboardSheet.getRange("C3:F3").format.font.size = 11;
			dashboardSheet.getRange("C3:F3").format.font.name = "Franklin Gothic Medium";
			dashboardSheet.getRange("C4:F8").format.font.name = "Bookman Old Style";
			dashboardSheet.getRange("C4:F7").format.font.size = 10;
			dashboardSheet.getRange("C3:F8").format.borders.getItem("InsideHorizontal").style = "Continuous";
			dashboardSheet.getRange("C3:F8").format.borders.getItem('EdgeBottom').style = 'Continuous';
			dashboardSheet.getRange("C3:F8").format.borders.getItem('EdgeTop').style = 'Continuous';
			dashboardSheet.getRange("C8:F8").format.font.size = 11;
			dashboardSheet.getRange("C8:F8").format.font.name = "Franklin Gothic Medium";


			// Queue commands to add the SEMESTER CREDITS and CLASSES section
			var semestersummarydata = [["SEMESTER", "CREDITS", "CLASSES"],
						   ["Semester 1", '=IFERROR(SUMIFS(D22:D122,F22:F122,C12),"")', '=IFERROR(COUNTIFS(F22:F122,C12),"")'],
						   ["Semester 2", '=IFERROR(SUMIFS(D22:D122,F22:F122,C13),"")', '=IFERROR(COUNTIFS(F22:F122,C13),"")'],
						   ["Semester 3", '=IFERROR(SUMIFS(D22:D122,F22:F122,C14),"")', '=IFERROR(COUNTIFS(F22:F122,C14),"")'],
						   ["Semester 4", '=IFERROR(SUMIFS(D22:D122,F22:F122,C15),"")', '=IFERROR(COUNTIFS(F22:F122,C15),"")'],
						   ["Semester 5", '=IFERROR(SUMIFS(D22:D122,F22:F122,C16),"")', '=IFERROR(COUNTIFS(F22:F122,C16),"")'],
						   ["Semester 6", '=IFERROR(SUMIFS(D22:D122,F22:F122,C17),"")', '=IFERROR(COUNTIFS(F22:F122,C17),"")'],
						   ["Semester 7", '=IFERROR(SUMIFS(D22:D122,F22:F122,C18),"")', '=IFERROR(COUNTIFS(F22:F122,C18),"")'],
						   ["Semester 8", '=IFERROR(SUMIFS(D22:D122,F22:F122,C19),"")', '=IFERROR(COUNTIFS(F22:F122,C19),"")'],
						   ["Total", "=sum(D12:D19)", "=sum(E12:E19)"]];
			dashboardSheet.getRange("C11:E20").values = semestersummarydata;

			// Queue commands to format the SEMESTER section
			dashboardSheet.getRange("C11:E11").format.font.size = 11;
			dashboardSheet.getRange("C11:E11").format.font.name = "Franklin Gothic Medium";
			dashboardSheet.getRange("C12:E19").format.font.name = "Bookman Old Style";
			dashboardSheet.getRange("C12:E19").format.font.size = 10;
			dashboardSheet.getRange("C11:E20").format.borders.getItem("InsideHorizontal").style = "Continuous";
			dashboardSheet.getRange("C11:E20").format.borders.getItem('EdgeBottom').style = 'Continuous';
			dashboardSheet.getRange("C11:E20").format.borders.getItem('EdgeTop').style = 'Continuous';
			dashboardSheet.getRange("C20:E20").format.font.size = 11;
			dashboardSheet.getRange("C20:E20").format.font.name = "Franklin Gothic Medium";
			dashboardSheet.getRange("C20:E20").format.font.color = "white";
			dashboardSheet.getRange("C20:E20").format.fill.color = "2A4C69";


			// Queue commands to create the chart based on the semester data
			// Queue commands to set the title for the chart
			var charttitle = "SEMESTER SUMMARY";
			dashboardSheet.getRange("A3:A3").values = charttitle;
			dashboardSheet.getRange("A3:A3").format.font.name = "Franklin Gothic Medium";
			dashboardSheet.getRange("A3:A3").format.font.size = 11;
			var semestersummarytitle = "Semester Summary Data";
			dashboardSheet.getRange("C10:C10").values = semestersummarytitle;
			dashboardSheet.getRange("C10:C10").format.font.name = "Century";
			dashboardSheet.getRange("C10:E10").format.font.size = 11;
			dashboardSheet.getRange("C10:E10").format.fill.color = "1E8FEB";
			dashboardSheet.getRange("C10:E10").format.font.color = "white";

			// Queue commands to create a chart and format it
			var chartDataRange = dashboardSheet.getRange("C11:E19");
			var semestersummarychart = dashboardSheet.charts.add("BarClustered", chartDataRange, "auto");
			semestersummarychart.setPosition("A4", "A19");
			semestersummarychart.legend.position = "right";
			semestersummarychart.title.visible = false;
			semestersummarychart.dataLabels.showValue = true;
			semestersummarychart.series.getItemAt(0).format.fill.setSolidColor("green");

			// Run the queued-up commands, and return a promise to indicate task completion
			return ctx.sync();

		})
		.catch(function (error) {
			// Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
			app.showNotification("Error: " + error);
			console.log("Error: " + error);
			if (error instanceof OfficeExtension.Error) {
				console.log("Debug info: " + JSON.stringify(error.debugInfo));
			}
		});
	}

	// Adds a course to the Courses table
	function addCourseToTable() {

		// Run a batch operation against the Excel object model
		Excel.run(function (ctx) {

			// Create a proxy for the tables rows
			var tableRows = ctx.workbook.tables.getItem('CoursesTable').rows;

			// Queue a command to add a new row at the end of the table for the course
			tableRows.add(null, [[$("#course-name option:selected").text(), $("#course-id").val(), $("#credit-type").val(), $("#credits").val(), $("input[type='radio']:checked").val(), $("#semester").val()]]);

			// Run the queued-up commands, and return a promise to indicate task completion
			return ctx.sync();
		})
		.catch(function (error) {
			// Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
			app.showNotification("Error: " + error);
			console.log("Error: " + error);
			if (error instanceof OfficeExtension.Error) {
				console.log("Debug info: " + JSON.stringify(error.debugInfo));
			}
		});
	}
})();

/* 
Excel-Add-in-JS-CollegeCreditsTracker, https://github.com/OfficeDev/Excel-Add-in-JS-CollegeCreditsTracker

Copyright (c) Microsoft Corporation

All rights reserved.

MIT License:

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
associated documentation files (the "Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the
following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial
portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT
LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT
SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE
USE OR OTHER DEALINGS IN THE SOFTWARE.
*/