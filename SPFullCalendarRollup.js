<!--/*
 * SPFullCalendarRollup - Create a calendar view using FullCalendar.io and SharePoint Online Task list
 * Version 1.0 
 * @requires jQuery v1.11 or greater 
 * @requires jQuery, FullCalendar.io, Moment.js 
 *
 * Copyright (c) 2017 Mark Rackley / PAIT Group
 * Licensed under the MIT license:
 * http://www.opensource.org/licenses/mit-license.php
 */
/**
 * @description Create a calendar view using FullCalendar.io and SharePoint Online Task list
 * @type jQuery
 * @name SPFullCalendarRollup
 * @category Plugins/SPFullCalendarRollup
 * @author Mark Rackley / http://www.paitgroup.com / mrackley@paitgroup.com
 * 
 * for code to work "as is":
 * Map ows_StartDate and ows_DueDate to RefinableDate00
 * Map ows_DueDate to RefinableDate01
 * Map ows_StartDate to RefinableDaet02
 * 
 */
-->
<script type="text/javascript" src="//code.jquery.com/jquery-1.11.1.min.js"></script> 
<script  type="text/javascript" src="//cdnjs.cloudflare.com/ajax/libs/moment.js/2.10.6/moment.min.js"></script>

<script  type="text/javascript" src="//cdnjs.cloudflare.com/ajax/libs/fullcalendar/3.4.0/fullcalendar.min.js"></script>
<link  type="text/css" rel="stylesheet" href="//cdnjs.cloudflare.com/ajax/libs/fullcalendar/3.4.0/fullcalendar.min.css" /> 

<div id='calendar'></div>

<script type="text/javascript">

  var PATH_TO_SITE = "https://mytenant.sharepoint.com/sites/stuff";
  var TASK_LIST = "Tasks";
  var COLORS = ['#466365', '#B49A67', '#93B7BE', '#E07A5F', '#849483', '#084C61', '#DB3A34'];  

   DisplayTasks();
   
   function DisplayTasks()
   {
  	$('#calendar').fullCalendar( 'destroy' );
    $('#calendar').fullCalendar({

        header: {
            left: 'prev,next today',
            center: 'title',
            right: 'month,basicWeek,basicDay'
        },
        timezone: "UTC",
        //put the events on the calendar 
        events: function (start, end, timezone, callback) {
            startDate = start.format('YYYY-MM-DD');
            endDate = end.format('YYYY-MM-DD');
			
			var RESTQuery = "/_api/search/query\
					?querytext='SPContenttype:Task PATH:"+PATH_TO_SITE+"((RefinableDate00>"+startDate+" AND RefinableDate00<"+endDate+")OR\
					(RefinableDate00<"+startDate+" AND RefinableDate00>"+endDate+"))\
					'&selectproperties='SiteTitle,Title,Url,AssignedTo,ID,RefinableDate02,RefinableDate01'&rowlimit=250";
			
			var opencall = $.ajax({
		    		url: _spPageContextInfo.webAbsoluteUrl + RESTQuery,
		    		type: "GET",
		    		dataType: "json",
		    		headers: {
		    			Accept: "application/json;odata=verbose"
		    		}
		    });

            opencall.done(function (data, textStatus, jqXHR) {
			    	var events = [];
					var siteColors = {};
            		var colorNo = 0;
					
                    var rows = data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;
                    for (index in rows)
                    {
                        var cells = rows[index].Cells.results;
                        var title = "";
                        var assignedTo = "";
                        var start = "";
                        var due = "";
                        var url = "";
                        var id = "";
                        var site = "";
                        for (index2 in cells)
                        {
                            var cell = cells[index2];
                            console.log("------------" + cell.Key);
                            switch(cell.Key) {
                                case "AssignedTo":
                                    assignedTo = cell.Value;
                                    break;
                                case "Title":
                                    title = cell.Value;
                                    break;
                                case "Url":
                                    url = cell.Value;
                                    break;
                                case "RefinableDate02":
                                    start = cell.Value;
                                    break;
                                case "RefinableDate01":
                                    due = cell.Value;
                                    break;
                                case "ID":
                                    id = cell.ID;
                                    break;
                                case "SiteTitle":
                                    site = cell.Value;
                                    break;									
                                default:
                                    break;
                            }
                        }

						var color = siteColors[site];
						if (!color) {
							color = COLORS[colorNo++];
							siteColors[site] = color;
						}
						if (colorNo >= COLORS.length) {
							colorNo = 0;
						}

						events.push({
									title: site + ": " + title + " - " + assignedTo,
									id: id,
									start: start,
									end: due,
									url: url,
									color: color // specify the background color and border color can also create a class and use className paramter. 
								});
		    		}
					
					callback(events);
            
		    });
		}
	});
}


</script>
