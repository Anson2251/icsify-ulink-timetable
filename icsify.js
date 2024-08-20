// ==UserScript==
// @name         ICSify Ulink iSAMS Timetable
// @namespace    http://tampermonkey.net/
// @version      2024-08-19
// @description  Export Ulink iSAMS Timetable to an ICS File
// @author       Heyan Zhu
// @match        https://pupils.ulinkedu.com/api/profile/timetable/
// @icon         https://pupils.ulinkedu.com/styling/icons/head/favicon-32x32.png
// @grant        none
// ==/UserScript==

(async function() {
	"use strict";
	
	/* 	ics.js 
		A browser friendly .ics/.vcs file generator written entirely in JavaScript! 
	   
		MIT License

		Copyright (c) 2018 Travis Krause

		Permission is hereby granted, free of charge, to any person obtaining a copy
		of this software and associated documentation files (the "Software"), to deal
		in the Software without restriction, including without limitation the rights
		to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
		copies of the Software, and to permit persons to whom the Software is
		furnished to do so, subject to the following conditions:

		The above copyright notice and this permission notice shall be included in all
		copies or substantial portions of the Software.

		THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
		IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
		FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
		AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
		LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
		OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
		SOFTWARE.

	*/

	/* global saveAs, Blob, BlobBuilder, console */
	/* exported ics */
	var ics = function(uidDomain, prodId) {
		"use strict";

		if (
			navigator.userAgent.indexOf("MSIE") > -1 &&
			navigator.userAgent.indexOf("MSIE 10") == -1
		) {
			console.log("Unsupported Browser");
			return;
		}

		if (typeof uidDomain === "undefined") {
			uidDomain = "default";
		}
		if (typeof prodId === "undefined") {
			prodId = "Calendar";
		}

		var SEPARATOR = navigator.appVersion.indexOf("Win") !== -1 ? "\r\n" : "\n";
		var calendarEvents = [];
		var calendarStart = [
			"BEGIN:VCALENDAR",
			"PRODID:" + prodId,
			"VERSION:2.0",
		].join(SEPARATOR);
		var calendarEnd = SEPARATOR + "END:VCALENDAR";
		var BYDAY_VALUES = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"];

		return {
			/**
			 * Returns events array
			 * @return {array} Events
			 */
			events: function() {
				return calendarEvents;
			},

			/**
			 * Returns calendar
			 * @return {string} Calendar in iCalendar format
			 */
			calendar: function() {
				return (
					calendarStart +
					SEPARATOR +
					calendarEvents.join(SEPARATOR) +
					calendarEnd
				);
			},

			/**
			 * Add event to the calendar
			 * @param  {string} subject     Subject/Title of event
			 * @param  {string} description Description of event
			 * @param  {string} location    Location of event
			 * @param  {string} begin       Beginning date of event
			 * @param  {string} stop        Ending date of event
			 */
			addEvent: function(subject, description, location, begin, stop, rrule) {
				// I'm not in the mood to make these optional... So they are all required
				if (
					typeof subject === "undefined" ||
					typeof description === "undefined" ||
					typeof location === "undefined" ||
					typeof begin === "undefined" ||
					typeof stop === "undefined"
				) {
					return false;
				}

				// validate rrule
				if (rrule) {
					if (!rrule.rrule) {
						if (
							rrule.freq !== "YEARLY" &&
							rrule.freq !== "MONTHLY" &&
							rrule.freq !== "WEEKLY" &&
							rrule.freq !== "DAILY"
						) {
							throw "Recurrence rrule frequency must be provided and be one of the following: 'YEARLY', 'MONTHLY', 'WEEKLY', or 'DAILY'";
						}

						if (rrule.until) {
							if (isNaN(Date.parse(rrule.until))) {
								throw "Recurrence rrule 'until' must be a valid date string";
							}
						}

						if (rrule.interval) {
							if (isNaN(parseInt(rrule.interval))) {
								throw "Recurrence rrule 'interval' must be an integer";
							}
						}

						if (rrule.count) {
							if (isNaN(parseInt(rrule.count))) {
								throw "Recurrence rrule 'count' must be an integer";
							}
						}

						if (typeof rrule.byday !== "undefined") {
							if (
								Object.prototype.toString.call(rrule.byday) !== "[object Array]"
							) {
								throw "Recurrence rrule 'byday' must be an array";
							}

							if (rrule.byday.length > 7) {
								throw "Recurrence rrule 'byday' array must not be longer than the 7 days in a week";
							}

							// Filter any possible repeats
							rrule.byday = rrule.byday.filter(function(elem, pos) {
								return rrule.byday.indexOf(elem) == pos;
							});

							for (var d in rrule.byday) {
								if (BYDAY_VALUES.indexOf(rrule.byday[d]) < 0) {
									throw "Recurrence rrule 'byday' values must include only the following: 'SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'";
								}
							}
						}
					}
				}

				//TODO add time and time zone? use moment to format?
				var start_date = new Date(begin);
				var end_date = new Date(stop);
				var now_date = new Date();

				var start_year = ("0000" + start_date.getFullYear().toString()).slice(
					-4
				);
				var start_month = ("00" + (start_date.getMonth() + 1).toString()).slice(
					-2
				);
				var start_day = ("00" + start_date.getDate().toString()).slice(-2);
				var start_hours = ("00" + start_date.getHours().toString()).slice(-2);
				var start_minutes = ("00" + start_date.getMinutes().toString()).slice(
					-2
				);
				var start_seconds = ("00" + start_date.getSeconds().toString()).slice(
					-2
				);

				var end_year = ("0000" + end_date.getFullYear().toString()).slice(-4);
				var end_month = ("00" + (end_date.getMonth() + 1).toString()).slice(-2);
				var end_day = ("00" + end_date.getDate().toString()).slice(-2);
				var end_hours = ("00" + end_date.getHours().toString()).slice(-2);
				var end_minutes = ("00" + end_date.getMinutes().toString()).slice(-2);
				var end_seconds = ("00" + end_date.getSeconds().toString()).slice(-2);

				var now_year = ("0000" + now_date.getFullYear().toString()).slice(-4);
				var now_month = ("00" + (now_date.getMonth() + 1).toString()).slice(-2);
				var now_day = ("00" + now_date.getDate().toString()).slice(-2);
				var now_hours = ("00" + now_date.getHours().toString()).slice(-2);
				var now_minutes = ("00" + now_date.getMinutes().toString()).slice(-2);
				var now_seconds = ("00" + now_date.getSeconds().toString()).slice(-2);

				// Since some calendars don't add 0 second events, we need to remove time if there is none...
				var start_time = "";
				var end_time = "";
				if (
					start_hours +
					start_minutes +
					start_seconds +
					end_hours +
					end_minutes +
					end_seconds !=
					0
				) {
					start_time = "T" + start_hours + start_minutes + start_seconds;
					end_time = "T" + end_hours + end_minutes + end_seconds;
				}
				var now_time = "T" + now_hours + now_minutes + now_seconds;

				var start = start_year + start_month + start_day + start_time;
				var end = end_year + end_month + end_day + end_time;
				var now = now_year + now_month + now_day + now_time;

				// recurrence rrule vars
				var rruleString;
				if (rrule) {
					if (rrule.rrule) {
						rruleString = rrule.rrule;
					} else {
						rruleString = "RRULE:FREQ=" + rrule.freq;

						if (rrule.until) {
							var uDate = new Date(Date.parse(rrule.until)).toISOString();
							rruleString +=
								";UNTIL=" +
								uDate.substring(0, uDate.length - 13).replace(/[-]/g, "") +
								"000000Z";
						}

						if (rrule.interval) {
							rruleString += ";INTERVAL=" + rrule.interval;
						}

						if (rrule.count) {
							rruleString += ";COUNT=" + rrule.count;
						}

						if (rrule.byday && rrule.byday.length > 0) {
							rruleString += ";BYDAY=" + rrule.byday.join(",");
						}
					}
				}

				var stamp = new Date().toISOString();

				var calendarEvent = [
					"BEGIN:VEVENT",
					"UID:" + calendarEvents.length + "@" + uidDomain,
					"CLASS:PUBLIC",
					"DESCRIPTION:" + description,
					"DTSTAMP;VALUE=DATE-TIME:" + now,
					"DTSTART;VALUE=DATE-TIME:" + start,
					"DTEND;VALUE=DATE-TIME:" + end,
					"LOCATION:" + location,
					"SUMMARY;LANGUAGE=en-us:" + subject,
					"TRANSP:TRANSPARENT",
					"END:VEVENT",
				];

				if (rruleString) {
					calendarEvent.splice(4, 0, rruleString);
				}

				calendarEvent = calendarEvent.join(SEPARATOR);

				calendarEvents.push(calendarEvent);
				return calendarEvent;
			},

			/**
			 * Download calendar using the saveAs function from filesave.js
			 * @param  {string} filename Filename
			 * @param  {string} ext      Extention
			 */
			download: function(filename, ext) {
				if (calendarEvents.length < 1) {
					return false;
				}

				ext = typeof ext !== "undefined" ? ext : ".ics";
				filename = typeof filename !== "undefined" ? filename : "calendar";
				var calendar =
					calendarStart +
					SEPARATOR +
					calendarEvents.join(SEPARATOR) +
					calendarEnd;

				var blob;
				if (navigator.userAgent.indexOf("MSIE 10") === -1) {
					// chrome or firefox
					blob = new Blob([calendar]);
				} else {
					// ie
					var bb = new BlobBuilder();
					bb.append(calendar);
					blob = bb.getBlob(
						"text/x-vCalendar;charset=" + document.characterSet
					);
				}
				saveAs(blob, filename + ext);
				return calendar;
			},

			/**
			 * Build and return the ical contents
			 */
			build: function() {
				if (calendarEvents.length < 1) {
					return false;
				}

				var calendar =
					calendarStart +
					SEPARATOR +
					calendarEvents.join(SEPARATOR) +
					calendarEnd;

				return calendar;
			},
		};
	};

	function addScriptDep(url) {
		return new Promise((resolve) => {
			let script = document.createElement("script");
			script.src = url;
			script.type = "text/javascript";
			script.async = true;
			script.onload = () => {
				console.log("[ICSify Script] Successfully loaded dependency" + url);
				resolve();
			};
			document.head.appendChild(script);
		});
	}

	function processCell(cell) {
		const data = cell.innerText
			.trim()
			.split("\n")
			.map((s) => (s || "").trim());
		return {
			lesson: data[0] || "",
			teacher: data[1] || "",
			room: (data[2] !== "Null" ? data[2] : null) || "",
		};
	}

	function formatDate(date) {
		let year = date.getFullYear();
		let month = (date.getMonth() + 1).toString().padStart(2, "0"); // Months are 0-indexed
		let day = date.getDate().toString().padStart(2, "0");

		return `${year}-${month}-${day}`;
	}

	function generateICS(timeTable, startDate, endDate) {
		let cal = ics();
		for (let i = 0; i < timeTable.length; i++) {
			for (let j = 0; j < 5; j++) {
				const cell = timeTable[i][j];
				if (!cell.lesson) continue;
				let startDate1 = new Date(
					`${formatDate(startDate)}T${cell.time.start}`
				);
				let endDate1 = new Date(`${formatDate(startDate)}T${cell.time.end}`);

				startDate1.setDate(startDate1.getDate() + j);
				endDate1.setDate(endDate1.getDate() + j);

				cal.addEvent(
					cell.lesson,
					`with ${cell.teacher}`,
					cell.room,
					startDate1.toISOString(),
					endDate1.toISOString(), {
						freq: "WEEKLY",
						until: endDate.toISOString(),
					}
				);
			}
		}

		return cal;
	}

	function getTimeTable() {
		const lessionDurnation = 45;
		const table = document.querySelector(".TTB_Table");

		return [...table.tBodies[0].childNodes].slice(2).map((row) => {
			let time = row.childNodes[0].innerText.split("\n")[1].trim();
			return [...row.childNodes].slice(1).map((cell) => {
				const details = processCell(cell);
				details.time = {
					// hh:mm:ss
					start: time + ":00",
					end: (() => {
						let [hour, min] = time.split(":").map((n) => parseInt(n));
						min += lessionDurnation;
						hour += Math.floor(min / 60);
						min = min % 60;
						return `${String(hour).padStart(2, "0")}:${String(min).padStart(
              2,
              "0"
            )}:00`;
					})(),
				};
				return details;
			});
		});
	}

	function askDateStartEnd() {
		return new Promise((resolve) => {
			const date = new Date(
				prompt(
					"Please enter a start date for the time table (YYYY-MM-DD)",
					new Date().toISOString().split("T")[0]
				)
			);
			const endDate = new Date(
				prompt(
					"Please enter an end date for the time table (YYYY-MM-DD)",
					(() => {
						let d = new Date();
						d.setDate(d.getDate() + 7);
						return d.toISOString().split("T")[0];
					})()
				)
			);
			resolve([date, endDate]);
		});
	}

	async function icsify() {
		let [start, end] = await askDateStartEnd();

		const timeTable = getTimeTable();
		generateICS(timeTable, start, end).download();
	}

	window.icsify = icsify;

	await addScriptDep(
		"https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js"
	);

	document.querySelector(
		".SP_Page_Content > table:nth-child(2) > tbody:nth-child(1) > tr:nth-child(1)"
	).innerHTML += `<td><button onclick="icsify()">Export to ICS</button></td>`;
})();