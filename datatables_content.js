var debug = false;

function log(message) {
  if (debug) {
    log(message);
  }
}

dataSet = new Array();

// Load neccessary SharePoint modules
$(document).ready(function() {
  var scriptbase = _spPageContextInfo.webServerRelativeUrl + "/_layouts/15/";

  $.getScript(scriptbase + "SP.Runtime.js", function() {
    $.getScript(scriptbase + "SP.js", function() {
      $.getScript(scriptbase + "SP.Taxonomy.js", execOperation);
    });
  });
});

// Main operations
function execOperation() {
  /* SP Utils */
  function getAppAbsoluteUrl() {
    return _spPageContextInfo.webAbsoluteUrl;
  }

  function getAppRelativeUrl() {
    return _spPageContextInfo.webServerRelativeUrl;
  }

  function getAppSiteCollectionUrl() {
    return _spPageContextInfo.siteAbsoluteUrl;
  }

  function getAppODataApiUrl() {
    return getAppAbsoluteUrl() + "/_api";
  }

  var baseRequest = {
    url: "",
    type: ""
  };

  /* Create a new OData request for JSON response */
  function getRequest(endpoint) {
    var request = baseRequest;
    request.type = "GET";
    request.url = endpoint;
    request.headers = { ACCEPT: "application/json;odata=verbose" };

    return request;
  }

  function onFailure() {
    $(window).load(function() {
      function show_error_message() {
        $("#err_msg").html(
          "<div class='alert alert-danger' role='alert'><strong>Content Contection Error</strong><p>Sorry, we are unable to get the content at this time.  Please contact support.</p></div>"
        );
      }
      window.setTimeout(show_error_message, 10);
    });
  }

  function newsArticles() {
    // get all site links
    var newsList = "News Articles";
    var htmlContainer = "#news-container";
    var htmlSection = "#news-content";
    var htmlErr = "#news-err";
    var query =
      getAppODataApiUrl() +
      "/web/lists/getbytitle('" +
      newsList +
      "')/Items" +
      "?$select=Title,FileLeafRef,DocIcon,OData__dlc_DocIdUrl,ArticleStatus,AKNDescription&$top=1000" +
      "&$filter=(ArticleStatus eq 'Active') or (ArticleStatus eq 'Inactive')" +
      "&$orderby=Modified desc";

    // execute query
    $.ajax(getRequest(query))
      .then(function onSuccess(data) {
        itemCount = data.d.results.length;
        if (itemCount <= 0) {
          $(htmlContainer).hide(); // No Results, so don't display.
          log("No news articles found. Hiding news articles.");
        } else {
          htmlStr = "";

          $.each(data.d.results, function(index, item) {
            var title = item.Title;
            var linkUrl = item.OData__dlc_DocIdUrl.Url;
            var descriptionVal = item.AKNDescription;
            var audienceVal = "Systemwide";
            var displayName = "";

            htmlStr += "<p><a href='" + linkUrl + "'>" + title + "</a><br>";
            if (descriptionVal) {
              htmlStr += descriptionVal + "<br>";
            }

            htmlStr += audienceVal + "</p>";
            htmlStr += "<br>";

            $(htmlSection).html(htmlStr);
            $(htmlSection).show();
            $("#accordion").hide();
            log("Displaying news articles.");
          });
        }
      })
      .fail(onFailure);
  }

  function contacts() {
    // get all site links
    var contactsList = "Contacts";
    var htmlContainer = "#contacts-container";
    var htmlSection = "#contacts-content";
    var htmlErr = "#contacts-err";
    var query =
      getAppODataApiUrl() +
      "/web/lists/getbytitle('" +
      contactsList +
      "')/Items" +
      "?$select=ContactType,Employee/Name,Employee/Title,Employee/EMail,Employee/WorkPhone,Employee/FirstName,Employee/LastName,Employee/JobTitle" +
      "&$expand=Employee" +
      "&$filter=ContactType ne 'Inactive'" +
      "&$orderby=ContactType";

    // execute query
    $.ajax(getRequest(query))
      .then(function onSuccess(data) {
        itemCount = data.d.results.length;
        if (itemCount <= 0) {
          $(htmlContainer).hide(); // No Results, so don't display.
          log("No contacts found. Hiding Contacts.");
        } else {
          htmlStr = "";

          $.each(data.d.results, function(index, item) {
            var firstName = item.Employee.FirstName;
            var lastName = item.Employee.LastName;
            var fullName = firstName + " " + lastName;
            var email = item.Employee.EMail;
            var jobTitle = item.Employee.JobTitle;
            var phone = item.Employee.WorkPhone;
            var empTitle = item.Employee.Title;
            var displayName = "";

            if (firstName) {
              displayName = fullName;
            } else {
              displayName = empTitle;
            }

            htmlStr += "<div class='well' class='contact-card'>";
            htmlStr += "<ul class='list-unstyled'>";
            htmlStr += "<li><strong>" + displayName + "</strong></li>";
            if (jobTitle) {
              htmlStr += "<li>" + jobTitle + "</li>";
            }
            if (phone) {
              htmlStr += "<li>" + phone + "</li>";
            }
            if (email) {
              htmlStr += "<li><a href='" + email + "'>" + email + "</a></li>";
            }
            htmlStr += "</ul>";
            htmlStr += "</div>";

            $(htmlSection).html(htmlStr);
            $(htmlSection).show();
            log("Displaying Contacts.");
          });
        }
      })
      .fail(onFailure);
  }

  function relatedlinks() {
    // get all related links
    var siteLinkList = "Related Links";
    var htmlContainer = "#relatedlinks-container";
    var htmlSection = "#related-links";
    var htmlErr = "#relatedlinks-err";
    var maxResults = 4;
    var maxItemsPerRow = 2;
    var query =
      getAppODataApiUrl() +
      "/web/lists/getbytitle('" +
      siteLinkList +
      "')/Items" +
      "?$select=*" +
      "&$top=" +
      maxResults +
      "&$orderby=Title";

    // execute query
    $.ajax(getRequest(query))
      .then(function onSuccess(data) {
        itemCount = data.d.results.length;
        if (itemCount <= 0) {
          $(htmlContainer).hide(); // No Results, so don't display.
          log("No related links found. Hiding RelatedLinks.");
        } else {
          var htmlStr = "";
          $.each(data.d.results, function(index, item) {
            var linkURL = item.SiteURL;
            var linkDisplayText = item.Title;
            var linkDescription = "";
            if (item.AKNDescription) {
              linkDescription = "<p>" + $(item.AKNDescription).text() + "</p>";
            }
            htmlStr +=
              '<div class="col-md-6"><a class="btn btn-default" role="button" href="' +
              linkURL +
              '" target="_blank">' +
              linkDisplayText +
              " " +
              linkDescription +
              "</a></div>";
            log("Related link: " + linkDisplayText + "(" + linkURL + ") added");
          });
          $(htmlSection).append(htmlStr);
          $(htmlContainer).show();
        }
      })
      .fail(onFailure);
  }

  function quicklinks() {
    // get all site links
    var siteLinkList = "Quick Links";
    var htmlSection = "#quick-links";
    var maxResults = 6;
    var maxItemsPerRow = 3;
    var query =
      getAppODataApiUrl() +
      "/web/lists/getbytitle('" +
      siteLinkList +
      "')/Items" +
      "?$select=*" +
      "&$top=" +
      maxResults +
      "&$orderby=Title";

    // execute query
    $.ajax(getRequest(query))
      .then(function onSuccess(data) {
        itemCount = data.d.results.length;
        if (itemCount <= 0) {
          $("#quicklinks-container").hide(); // No Results, so don't display.
          log("No quicklinks found. Hiding QuickLinks.");
        } else {
          var htmlStr = "";
          $.each(data.d.results, function(index, item) {
            var linkURL = item.SiteURL;
            var linkDisplayText = item.Title;
            var linkDescription = item.AKNDescription;
            if (index % maxItemsPerRow === 0) {
              htmlStr += '<div class="row">';
            }
            htmlStr += '<div class="col-md-4">';
            htmlStr +=
              '<a class="btn btn-default" role="button" href="' +
              linkURL +
              '" target="_blank"> ' +
              linkDisplayText +
              " &#160;";
            htmlStr += '<i class="fa fa-external-link" aria-hidden="true"></i>';
            if (linkDescription) {
              htmlStr +=
                '<div id="akn-quick-links-copy">' + linkDescription + "</div>";
            }
            htmlStr += "</a></div>";
            if ((index - (maxItemsPerRow - 1)) % maxItemsPerRow === 0) {
              htmlStr += "</div>";
            }

            log("Quick link: " + linkDisplayText + "(" + linkURL + ") added");
          });
          $(htmlSection).append(htmlStr);
          $("#quicklinks-container").show();
        }
      })
      .fail(onFailure);
  }

  function getLinks() {
    // get all site links
    var siteLinkList = "Site Links";
    var query =
      getAppODataApiUrl() +
      "/web/lists/getbytitle('" +
      siteLinkList +
      "')/Items" +
      "?$select=*" +
      "&$orderby=Title";

    // execute query
    $.ajax(getRequest(query))
      .then(function onSuccess(data) {
        var htmlStr = "";
        $.each(data.d.results, function(index, item) {
          if (index < 0) {
            log("No site links found");
          } else {
            // Check for null subtopics
            if (item.Subtopic) {
              sectionName = item.Subtopic.TermGuid;
            } else {
              if (item.Topic) {
                sectionName = item.Topic.TermGuid + "Root";
              }
            }

            var linkURL = item.URL.Url;
            var linkDisplayText = item.URL.Description;

            var htmlSection = "#" + sectionName;
            htmlStr =
              '<p><a href="' +
              linkURL +
              '" target="_blank"><i class="fa fa-external-link" aria-hidden="true"></i> ' +
              linkDisplayText +
              "</a></p>";
            $(htmlSection).append(htmlStr);
            log("Site link" + linkDisplayText + "(" + linkURL + ") added");
          }
        });
      })
      .fail(onFailure);
  }

  function getDocs(listName) {
    // get all documents
    var query =
      getAppODataApiUrl() +
      "/web/lists/getbytitle('" +
      listName +
      " Documents')/Items" +
      "?$select=*,FileLeafRef" +
      "&$orderby=Title";

    // execute query
    $.ajax(getRequest(query))
      .then(function onSuccess(data) {
        var htmlStr = "";
        $.each(data.d.results, function(index, item) {
          if (index < 0) {
            //htmlStr += "No Results"
          } else {
            // Check for null subtopics
            if (item.Subtopic) {
              sectionName = item.Subtopic.TermGuid;
            } else {
              if (item.Topic) {
                sectionName = item.Topic.TermGuid + "Root";
              }
            }

            var docTitle = item.Title;
            var docFileName = item.FileLeafRef;
            var docExtension = "";
            var fileTitle = "";
            var documentDisplayTitle = "";

            // Title and document extension information
            if (docFileName) {
              var n = docFileName.lastIndexOf(".");
              if (n > 0) {
                docExtension = docFileName.substring(n + 1, docFileName.length);
                fileTitle = docFileName.substring(0, n);
              }
              if (docTitle) {
                documentDisplayTitle = docTitle;
              } else {
                documentDisplayTitle = fileTitle;
              }
            }

            htmlSection = "#" + sectionName;
            htmlStr =
              '<p><a href="' +
              item.OData__dlc_DocIdUrl.Url +
              '" target="_blank"><img class="ms-rtePosition-4" src="/PublishingImages/' +
              docExtension +
              '.gif" alt="" style="width: 16px; height: 17px;"/> ' +
              documentDisplayTitle +
              "</a></p>";
            $(htmlSection).append(htmlStr);
          }
        });
      })
      .fail(onFailure);
  }

  /*
	function getContentItems(listName)
	{
        var docListFullName = listName + " Documents"
        var docListSystemName = docListFullName.substring(0,50);
        
        var linkListFullName = listName + " Links"
        var linkListSystemName = linkListFullName.substring(0,50);

        var docQuery = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/GetByTitle('" + docListSystemName + "')/items?$select=Title,Topic,Subtopic,FileLeafRef,DocIcon,OData__dlc_DocIdUrl&$top=1000";
		var linkQuery = _spPageContextInfo.webAbsoluteUrl + "/_api/web/lists/GetByTitle('" + linkListSystemName + "')/items?$select=Title,Topic,Subtopic,SiteURL&$top=1000";
		var siteGuid = "";
		var linkList = [];
		var htmlStr = "";

		// Get Docs
		$.ajax({
			url: docQuery,
			type: "GET",
			headers: {
			"accept": "application/json;odata=verbose",
			},
			success: function(data)
			{
				$.each(data.d.results, function(index, item)
				{
					var linkTitle = "";
					var linkUrl = item.OData__dlc_DocIdUrl.Url;
					var linkIcon = item.DocIcon;
					var linkSection = "";
					var linkTopic = item.Topic.TermGuid;
					var linkSubtopic = "";
					
					// Only set the subtopic if it exists
					if(item.Subtopic){linkSubtopic = item.Subtopic.TermGuid;}
					
					// Set the lowest level for AKNSource
					if(linkSubtopic)
					{
						linkSection = "#" + linkSubtopic;
					}
					else
					{
						linkSection = "#" + linkTopic;
					}
					
					// Title vs. Filename
					var docTitle = item.Title;
					var docFileName = item.FileLeafRef;
					var fileTitle = "";
					
					// Title and document extension information
					if(docFileName)
					{
						var n = docFileName.lastIndexOf(".");
						if (n > 0) 
						{
							fileTitle = docFileName.substring(0,n);
						}
						if(docTitle)
						{
							linkTitle = docTitle;
						}
						else
						{
							linkTitle = fileTitle;
						}
					}
					
					// Store documents in link list
					linkList.push({title: linkTitle, url: linkUrl, extension: linkIcon, section: linkSection});
				});
					// Get Links
					$.ajax({
						url: linkQuery,
						type: "GET",
						headers: {
						"accept": "application/json;odata=verbose",
						},
						success: function(data)
						{
							$.each(data.d.results, function(index, item)
							{
								//var linkTitle = item.Link.Description;
								//var linkUrl = item.Link.Url;
								var linkTitle = item.Title;
								var linkUrl = item.SiteURL;
								var linkIcon = "link";
								var linkSection = "";
								var linkTopic = item.Topic.TermGuid;
								var linkSubtopic = "";
								
								// Only set the subtopic if it exists
								if(item.Subtopic){linkSubtopic = item.Subtopic.TermGuid;}

								// Set the lowest level for AKNSource
								if(linkSubtopic)
								{
									linkSection = "#" + linkSubtopic;
								}
								else
								{
									linkSection = "#" + linkTopic;
								}
								
								// Store links in link list
								linkList.push({title: linkTitle, url: linkUrl, extension: linkIcon, section: linkSection});
							});
							
							// Sort Link List
							linkList.sort(dynamicSort("title"));
							
							// Check it out
                            log(linkList);
                            
                            var scURL = getAppSiteCollectionUrl();
							
							// Presentation
							for (var i = 0; i < linkList.length; i++) 
							{
								
								var imageExt;
								var docExt = linkList[i].extension;

								switch(docExt)
								{
									case "doc": imageExt = "doc.gif"; break;
                                    case "docx": imageExt = "doc.gif"; break;
                                    case "docm": imageExt = "doc.gif"; break;

                                    case "xls": imageExt = "xls.gif"; break;
                                    case "xlsx": imageExt = "xls.gif"; break;
                                    case "xlsm": imageExt = "xls.gif"; break;

                                    case "ppt": imageExt = "ppt.gif"; break;
                                    case "pptx": imageExt = "ppt.gif"; break;
                                    case "pps": imageExt = "ppt.gif"; break;
                                    case "ppsx": imageExt = "ppt.gif"; break;
                                    
                                    case "vsd": imageExt = "vsd.png"; break;
                                    case "vsdx": imageExt = "vsd.png"; break;

                                    case "pdf": imageExt = "pdf.gif"; break;

                                    case "link": imageExt = "link.gif"; break;

                                    default: imageExt = "generic.gif";
								}
								
								htmlStr = '<p><a href="' + linkList[i].url + '" target="_blank"><img class="ms-rtePosition-4" src="' + scURL + '/PublishingImages/' + imageExt + '" alt="" style="width: 16px; height: 17px;"/> ' + linkList[i].title + '</a></p>';
								$(linkList[i].section).append(htmlStr);
							}
							
						},
						error: function(error)
						{
							log("Oops! An error occurred with the link load: " + JSON.stringify(error));
						}
					});	
			},
			error: function(error)
			{
				log("Oops! An error occurred with the document load: " + JSON.stringify(error));
			}
		});
	}
	*/

  function getContentItems(
    stopicGuid,
    stopicName,
    stName,
    sortMethod,
    loadSubtopics
  ) {
    //Share Point limits library titles to 50 characters.
    //Back-end scripts that build the site and libraries truncates the library title.
    //Do the same here to get proper library title.
    var docListFullName = stopicName + " Documents";
    var docListSystemName = docListFullName.substring(0, 50);

    var linkListFullName = stopicName + " Links";
    var linkListSystemName = linkListFullName.substring(0, 50);

    //var dataSet = new Array();

    //log("getContentItems(" + stopicGuid + "," + stopicName + "," + sortMethod + "," + (loadSubtopics?"true":"false") + ")");
    var includeSubtopics = false;
    if (loadSubtopics) {
      includeSubtopics = loadSubtopics;
    }
    var docQuery =
      _spPageContextInfo.webAbsoluteUrl +
      "/_api/web/lists/GetByTitle('" +
      docListSystemName +
      "')/items?$select=Title,Topic,Subtopic,FileLeafRef,DocIcon,OData__dlc_DocIdUrl,Modified" +
      "&$orderBy=Title";

    var linkQuery =
      _spPageContextInfo.webAbsoluteUrl +
      "/_api/web/lists/GetByTitle('" +
      linkListSystemName +
      "')/items?$select=Title,SiteURL,Topic,Subtopic,Modified" +
      "&$orderBy=Title";

    log(
      "getContentItems(" +
        stopicGuid +
        "," +
        stopicName +
        "," +
        stName +
        "," +
        sortMethod +
        "," +
        (loadSubtopics ? "true" : "false") +
        ") + Doc Query: " +
        docQuery
    );
    var siteGuid = "";
    var linkList = [];

    var htmlStr = "";
    var sortMethod = "";
    var byDateItems = [];
    var byAlphaItems = [];
    var uniqueNames = [];

    //alert("getContentItems: #" + stopicGuid);
    //Clear the previous links from the DOM
    //$("#" + stopicGuid).html("");

    $.ajax({
      url: docQuery,
      type: "GET",
      headers: {
        accept: "application/json;odata=verbose"
      },
      success: function(data) {
        $.each(data.d.results, function(index, item) {
          var linkTitle = "";
          var linkUrl = item.OData__dlc_DocIdUrl.Url;
          var linkIcon = item.DocIcon;
          var linkSection = "";
          var linkTopic = item.Topic.TermGuid;
          var linkSubtopicName = "";

          //var linkSubtopic = item.Subtopic.TermGuid;
          var linkSubtopic = null;
          if (item.Subtopic) {
            linkSubtopic = item.Subtopic.TermGuid;
            linkSubtopicName = stName;
          }
          if (
            (linkSubtopic && includeSubtopics) ||
            (!linkSubtopic && !includeSubtopics)
          ) {
            var linkCreated = item.Modified;
            var d = new Date(linkCreated);
            linkCreated =
              d.getMonth() + 1 + "/" + d.getDate() + "/" + d.getFullYear();
            //var linkLabel = item.Subtopic.Label;

            // Set the lowest level for AKNSource
            if (linkSubtopic) {
              //linkSection = "#parent_" + linkSubtopic;
              linkSection = "#" + linkSubtopic;
            } else {
              //linkSection = "#parent_" + linkTopic;
              linkSection = "#" + linkTopic;
            }

            // Title vs. Filename
            var docTitle = item.Title;
            var docFileName = item.FileLeafRef;
            var fileTitle = "";

            // Title and document extension information
            if (docFileName) {
              var n = docFileName.lastIndexOf(".");
              if (n > 0) {
                fileTitle = docFileName.substring(0, n);
              }
              if (docTitle) {
                linkTitle = docTitle;
              } else {
                linkTitle = fileTitle;
              }
            }
            // Store documents in link list
            linkList.push({
              title: linkTitle,
              url: linkUrl,
              extension: linkIcon,
              section: linkSubtopicName,
              created: linkCreated
            });
            dataSet.push(
              new Array(
                linkTitle,
                linkUrl,
                linkIcon,
                linkSubtopicName,
                linkCreated
              )
            );
          }
        });

        // Get Links
        $.ajax({
          url: linkQuery,
          type: "GET",
          headers: {
            accept: "application/json;odata=verbose"
          },
          success: function(data) {
            $.each(data.d.results, function(index, item) {
              //var linkTitle = item.Link.Description;
              //var linkUrl = item.Link.Url;
              var linkTitle = item.Title;
              var linkUrl = item.SiteURL;
              var linkIcon = "link";
              var linkSection = "";
              var linkTopic = item.Topic.TermGuid;
              var linkSubtopicName = "";

              //var linkSubtopic = item.Subtopic.TermGuid;
              var linkSubtopic = null;
              if (item.Subtopic) {
                linkSubtopic = item.Subtopic.TermGuid;
                linkSubtopicName = stName;
              }
              if (
                (linkSubtopic && includeSubtopics) ||
                (!linkSubtopic && !includeSubtopics)
              ) {
                var linkCreated = item.Created;
                //var linkLabel = item.Subtopic.Label;
                // Set the lowest level for AKNSource
                if (linkSubtopic) {
                  //linkSection = "#parent_" + linkSubtopic;
                  linkSection = "#" + linkSubtopic;
                } else {
                  //linkSection = "#parent_" + linkTopic;
                  linkSection = "#" + linkTopic;
                }
                linkList.push({
                  title: linkTitle,
                  url: linkUrl,
                  extension: linkIcon,
                  section: linkSubtopicName,
                  created: linkCreated
                });
                dataSet.push(
                  new Array(
                    linkTitle,
                    linkUrl,
                    linkIcon,
                    linkSubtopicName,
                    linkCreated
                  )
                );
              }
            });
            var trmElementName = stopicName.replace(/ /gi, "").toLowerCase();
            console.log("dataSet: " + trmElementName);
            console.log(dataSet);

            $("#" + trmElementName).DataTable({
              data: dataSet,
              columns: [
                {
                  title: "Name"
                },

                {
                  title: "Link",
                  render: function(data, type, row, meta) {
                    data = '<a href="' + data + '">View Item</a>';
                    return data;
                  }
                },
                { title: "Filetype" },
                { title: "Section" },
                { title: "Date" }
              ]
            });
            // Check it out
            //console.log(linkList);

            // Presentation
            for (var i = 0; i < linkList.length; i++) {
              var currSecGuid = linkList[i].section;
              var sectionGuid = currSecGuid.substring(8);
              htmlStr =
                '<p><a href="' +
                linkList[i].url +
                '" target="_blank" id="p_' +
                sectionGuid +
                i +
                '"><img class="ms-rtePosition-4" src="/PublishingImages/' +
                linkList[i].extension +
                '.gif" alt="" style="width: 16px; height: 17px;"/> ' +
                linkList[i].title +
                "</a></p>";
              $(linkList[i].section + " .load-msg").addClass("akn-hidden");
              $(linkList[i].section).append(htmlStr);
              /*
							$("#p_" + sectionGuid + i).click(function(){
								var itemDivCont = document.getElementById("parent_" + sectionGuid);
								var expandedDiv = itemDivCont.classList.contains('collapsed');
								if(expandedDiv!=true){
									$("#parent_" + sectionGuid).removeClass("collapsed");
									//alert("Document Clicked");
								}
							})
							*/
              if ($(linkList[i].section).parents(".akn-hidden").length) {
                $(linkList[i].section)
                  .parents(".akn-hidden")
                  .each(function() {
                    $(this).removeClass("akn-hidden");
                    //$(this).addClass("akn-show");
                  });
              }
            }
          },
          error: function(error) {
            log(
              "Oops! An error occurred with the link load: " +
                JSON.stringify(error)
            );
          }
        });
      },
      error: function(error) {
        log(
          "Oops! An error occurred with the document load: " +
            JSON.stringify(error)
        );
      }
    });
  }

  //REST request to Subtopic Sort list on root site to get the specified sort method per subtopic
  function getSortMethod(stpGuid, stpName, cbfn) {
    log(
      "getSortMethod(" +
        stpGuid +
        "," +
        stpName +
        "," +
        (cbfn ? "defined function" : "null") +
        ")"
    );
    //sortQuery = _spPageContextInfo.siteAbsoluteUrl + "/_api/web/lists/GetByTitle('Subtopic%20Sort')/items?$select=Subtopic,Sort_x0020_Method&$filter=TaxCatchAll/IdForTerm eq '" + stpGuid + "'";
    sortQuery =
      _spPageContextInfo.siteAbsoluteUrl +
      "/_api/web/lists/GetByTitle('Accordion%20Sort')/items?$select=TopicOrSubtopic,SortMethod&$filter=TaxCatchAll/IdForTerm eq '" +
      stpGuid +
      "'";
    var subtopicSortMethod = "";
    var nullVal = "Title asc";
    var newSubtopicTermLabel = "";
    var ret = "";
    //if(cbfn == null) {
    //$.ajaxSetup({async: false});
    //}
    $.ajax({
      url: sortQuery,
      type: "GET",
      headers: {
        accept: "application/json;odata=verbose"
      },
      success: function(data) {
        log(data);
        if (data.d.results.length == 0) {
          if (cbfn != null) {
            cbfn(nullVal);
          } else {
            log(
              "getSortMethod(" +
                stpGuid +
                "," +
                stpName +
                "," +
                (cbfn ? "defined function" : "null") +
                ") returning " +
                nullVal
            );
            ret = nullVal;
          }
        } else {
          $.each(data.d.results, function(index, item) {
            //Assign sort method to variable and send back to callback
            if (stpGuid) {
              //subtopicSortMethod = item.Sort_x0020_Method;
              subtopicSortMethod = item.SortMethod;
              //return sort method to callback func
              if (cbfn != null) {
                if (subtopicSortMethod) {
                  cbfn(subtopicSortMethod);
                }
              } else {
                if (subtopicSortMethod) {
                  log(
                    "getSortMethod(" +
                      stpGuid +
                      "," +
                      stpName +
                      "," +
                      (cbfn ? "defined function" : "null") +
                      ") returning " +
                      subtopicSortMethod
                  );
                  ret = subtopicSortMethod;
                }
              }
            }
          });
        }
      }
    });
    //if(cbfn == null) {
    //$.ajaxSetup({async: true});
    //return ret;
    //}
  }
  /*
	function dynamicSort(property) 
	{
		var sortOrder = 1;
		if(property[0] === "-") {
			sortOrder = -1;
			property = property.substr(1);
		}
		return function (a,b) {
			var result = (a[property] < b[property]) ? -1 : (a[property] > b[property]) ? 1 : 0;
			return result * sortOrder;
		}
	}
	*/

  SP.ClientContext.prototype.executeQuery = function() {
    var deferred = $.Deferred();
    this.executeQueryAsync(
      function() {
        deferred.resolve(arguments);
      },
      function() {
        deferred.reject(arguments);
      }
    );
    return deferred.promise();
  };

  function getSubtopicItemCount(sTopicGUID, sTopicName, cbfn) {
    //Share Point limits library titles to 50 characters.
    //Back-end scripts that build the site and libraries truncates the library title.
    //Do the same here to get proper library title.
    var docListFullName = sTopicName + " Documents";
    var docListSystemName = docListFullName.substring(0, 50);

    var linkListFullName = sTopicName + " Links";
    var linkListSystemName = linkListFullName.substring(0, 50);

    var docCountQuery =
      _spPageContextInfo.webAbsoluteUrl +
      "/_api/web/lists/GetByTitle('" +
      docListSystemName +
      "')/items?$select=Title&$filter=TaxCatchAll/IdForTerm eq '" +
      sTopicGUID +
      "'";
    var linkCountQuery =
      _spPageContextInfo.webAbsoluteUrl +
      "/_api/web/lists/GetByTitle('" +
      linkListSystemName +
      "')/items?$select=Title&$filter=TaxCatchAll/IdForTerm eq '" +
      sTopicGUID +
      "'";
    var itemCount = 0;
    log(
      "getSubtopicItemcount(" +
        sTopicGUID +
        "," +
        sTopicName +
        ", function) : " +
        docCountQuery
    );
    //get docs
    $.ajax({
      url: docCountQuery,
      type: "GET",
      headers: {
        accept: "application/json;odata=verbose"
      },
      success: function(data) {
        itemCount = data.d.results.length;
        log(sTopicName + " - ItemCount = " + itemCount);
        //return itemCount to callback func
        if (cbfn != null) {
          cbfn(itemCount, sTopicGUID);
        }
        /*
                $.each(data.d.results, function(index, item)
                 {
                    //for each item returned increase itemCount var by one.
                    itemCount++;
                 });
                    // Get Links
                    $.ajax({
                        url: linkCountQuery,
                        type: "GET",
                        headers: {
                        "accept": "application/json;odata=verbose",
                        },
                        success: function(data)
                        {
                            //$.each(data.d.results, function(index, item)
                            //{
                            //    //for each item returned increase itemCount var by one.
                            //    itemCount++;
                            //});
							 itemCount = data.d.results.length;
							 log(sTopicName + ' - ItemCount = ' + itemCount);
                             //return itemCount to callback func
                             if(cbfn!=null){
                                 cbfn(itemCount, sTopicGUID);
                             }
                            
                        },
                        error: function(error)
                        {
                            log("Oops! An error occurred with the link load: " + JSON.stringify(error));
                        }
                    });	
					*/
      },
      error: function(error) {
        log(
          "Oops! An error occurred with the document load: " +
            JSON.stringify(error)
        );
      }
    });
  }

  function subtopicClicked(self, subtopicName) {
    var elemID = self.parentElement.getAttribute("id");
    var guidID = elemID.substring(7);
    log("subtopicClicked(): Elem ID = " + elemID + ":GUID = " + guidID);
    var getParent = document.getElementById(elemID);
    //check if 'collapsed' exists in content classes attribute, use for condition
    var classChecker = getParent.classList.contains("collapsed");

    if (classChecker == true) {
      //get sort method
      getSortMethod(guidID, subtopicName, function(sortMethod) {
        var sMethod = sortMethod;
        //Now that we've retreived the sort method pass information on to get subtopic content
        getContentItems(guidID, subtopicName, "", sMethod, true);
      });

      //If the subtopic is already expanded, then remove content and collapse subtopic
    } else {
      //var myElem = "collapse" + guidID;
      //var myDiv = document.getElementById(myElem);
      //while(myDiv.nextElementSibling){
      //        myDiv.parentNode.removeChild(myDiv.nextElementSibling);
      //}
      //$("#" + guidID + " .loader").removeClass("akn-hidden");
      //$("#" + guidID + " .loader").addClass("akn-show");
      //$("#" + guidID).html("");
      //$("#" + guidID).html('<div class="lds-roller"><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div></div>');
      $("#" + guidID).html('<div class="load-msg">Retrieving Items...</div>');
    }
  }

  function getSubtopicData(termName, termGUID) {
    var ctx = SP.ClientContext.get_current();
    var session = SP.Taxonomy.TaxonomySession.getTaxonomySession(ctx);
    var termStore = session.getDefaultSiteCollectionTermStore();
    var dataSet = new Array();
    var parentTermId = termGUID;
    var parentTerm = termStore.getTerm(parentTermId);
    var terms = parentTerm.get_terms(); //load child terms
    var htmlText = "";

    ctx.load(terms);
    var promise = ctx.executeQuery();

    promise.done(function() {
      log(termName + " subtopic term loaded.");
    });

    promise.then(
      function(sArgs) {
        //sArgs[0] == success callback sender
        //sArgs[1] == success callback args

        var sectionID = "#" + termGUID;

        for (var i = 0; i < terms.get_count(); i++) {
          var term = terms.getItemAtIndex(i);
          var trmGUID = term.get_id();
          var trmName = term.get_name();
          var trmDesc = term.get_description();
          console.log("Topic: " + termName + " Subtopic:" + trmName);

          var accordID = "accordionSub-" + trmGUID;
          /*
				htmlText += '<div class="col-md-12 subpanel-heading accordion-toggle collapsed" aria-expanded="false" data-toggle="collapse" data-target="#collapse' + trmGUID + '" data-parent="#accordionSub">';
				        htmlText += '<h2 id="h2_' + trmGUID + '" class="akn-subtopic-header">' + trmName + '</h2>';
				        htmlText += '<p id="akn-subtopic-paragraph">' + trmDesc+ '</p>';
				        htmlText += '<hr class="akn-subtopic-hr"/>';
				        htmlText += '<div id="collapse' + trmGUID + '" class="panel-collapse collapse" aria-expanded="false" style="height:0px;">';
				        htmlText += '</div>';
                htmlText += '</div>';
				*/

          htmlText += '<div id="' + trmGUID + '_ee" class="panel akn-hidden">';
          //htmlText += '<div class="col-md-12 subpanel-heading accordion-toggle collapsed ' + trmGUID + '_ee akn-hidden" aria-expanded="false" data-toggle="collapse" data-target="#collapse' + trmGUID + '" data-parent="#accordionSub" id="parent_' + trmGUID + '">';
          //htmlText += '<div class="col-md-12 subpanel-heading accordion-toggle collapsed ' + trmGUID + '_ee akn-hidden" aria-expanded="false" data-toggle="collapse" data-target="#collapse' + trmGUID + '" data-parent="' + sectionID + '" id="parent_' + trmGUID + '">';
          htmlText +=
            '<div class="subpanel-heading accordion-toggle collapsed ' +
            trmGUID +
            '_ee" aria-expanded="false" data-toggle="collapse" data-target="#collapse' +
            trmGUID +
            '" data-parent="' +
            sectionID +
            '" id="parent_' +
            trmGUID +
            '">';
          htmlText +=
            '<h2 id="h2_' +
            trmGUID +
            '" class="akn-subtopic-header">' +
            trmName +
            "</h2>";
          htmlText += '<p class="akn-subtopic-paragraph">' + trmDesc + "</p>";
          htmlText += '<hr class="akn-subtopic-hr"/>';
          htmlText += "</div>";
          //htmlText += '<div id="collapse' + trmGUID + '" class="panel-collapse collapse in" aria-expanded="false" style="height:0px;">';
          htmlText +=
            '<div id="collapse' +
            trmGUID +
            '" class="panel-collapse collapse" aria-expanded="false">';
          htmlText += '<div id="' + trmGUID + '">';
          //htmlText += '<div class="lds-roller"><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div></div>';
          htmlText += '<div class="load-msg">Retrieving Items...</div>';
          htmlText += "</div>";
          htmlText += "</div>";
          htmlText += "</div>";
          getContentItems(termGUID, termName, trmName, "", true);
        }
        getContentItems(termGUID, termName, trmName, "", false);
        log("Adding HTML for Subtopics..." + terms.get_count() + " total");
        $(sectionID).append(htmlText);
        //getDocs(termName);
        //getContentItems(termName);
        /*
			getSortMethod(termGUID, termName, function(sortMethod){
				var sMethod = sortMethod;
				//Now that we've retreived the sort method pass information on to get subtopic content
				getContentItems(termGUID, termName, sMethod);
			});
			*/
        log(
          "Time to add event listeners Subtopics..." +
            terms.get_count() +
            " total"
        );

        //getContentItems(guidID, subtopicName, sMethod, true);
        /*
        //Now that document is loaded with subtopic headers, set event listeners on each subtopic header
        for (var h = 0; h < terms.get_count(); h++) {
          var eventTerm = terms.getItemAtIndex(h);
          var eventTrmGUID = eventTerm.get_id();
          if (eventTrmGUID != null) {
            var myEl = document.getElementById("h2_" + eventTrmGUID);
            if (myEl) {
              log("Found Element for Event Listener:" + "h2_" + eventTrmGUID);
              myEl.addEventListener(
                "click",
                function() {
                  try {
                    subtopicClicked(this, termName).bind(myEl);
                  } catch (e) {
                    log(e.message);
                  }
                },
                false
              );
            }
          }
		}
		*/

        //Get count of docs and links to determine visibility of the topic header in the accordion

        log("Time to unhide Subtopics..." + terms.get_count() + " total");

        for (var j = 0; j < terms.get_count(); j++) {
          var countTerm = terms.getItemAtIndex(j);
          var countTrmGUID = countTerm.get_id();
          var countTrmName = countTerm.get_name();
          log(
            "Checking count for: [" + countTrmGUID + ", " + countTrmName + "]"
          );
          //getSubtopicItemCount(countTrmGUID, countTrmName, function(itemCount, cbnGUID){
          getSubtopicItemCount(countTrmGUID, termName, function(
            itemCount,
            cbnGUID
          ) {
            //log("getSubtopicItemCount(" + countTrmGUID + ", " + countTrmName + ")");
            log("Count: " + itemCount);
            if (itemCount > 0) {
              log("Unhide..." + cbnGUID);
              /**/
              var topicDiv = "#" + termGUID + "_ee";
              log("Remove hidden tag: " + topicDiv);
              $(topicDiv).removeClass("akn-hidden");
              var subtopicDiv = "#" + cbnGUID + "_ee";
              //var subtopicDiv = "#" + cbnGUID;
              log("Remove hidden tag: " + subtopicDiv);
              $(subtopicDiv).removeClass("akn-hidden");
              /**/
            }
          });
        }
      },

      function(fArgs) {
        //fArgs[0] == fail callback sender
        //fArgs[1] == fail callback args.
        //in JSOM the callback args aren't used much -
        //the only useful one is probably the get_message()
        //on the fail callback
        var failmessage = fArgs[1].get_message();
        log(termName + " subtopic term not loaded. Error: " + failmessage);
      }
    );
  }

  function loadTopicLevelLinksDocs(trmGUID, trmName) {
    var tGUID = trmGUID;
    var tName = trmName;
    getSortMethod(trmGUID, trmName, function(sortMethod) {
      var sort = sortMethod;
      log(trmName + " Sort: " + sort);
      //Now that we've retreived the sort method pass information on to get subtopic content
      log(trmName + " - Loading Topic Level Items");
      getContentItems(tGUID, tName, trmName, sort, false);
    });
  }

  function getTopicData(termName, termGUID) {
    var ctx = SP.ClientContext.get_current();
    var session = SP.Taxonomy.TaxonomySession.getTaxonomySession(ctx);
    var termStore = session.getDefaultSiteCollectionTermStore();

    var parentTermId = termGUID;
    var parentTerm = termStore.getTerm(parentTermId);
    var terms = parentTerm.get_terms(); //load child terms
    var scURL = getAppSiteCollectionUrl();
    var htmlText = "";

    ctx.load(terms);
    var promise = ctx.executeQuery();

    promise.done(function() {
      log(termName + " topic term loaded.");
    });

    promise.then(
      function(sArgs) {
        //sArgs[0] == success callback sender
        //sArgs[1] == success callback args

        for (var i = 0; i < terms.get_count(); i++) {
          var term = terms.getItemAtIndex(i);
          var trmGUID = term.get_id();
          var trmName = term.get_name();
          var trmListFullName = trmName.replace(/-|,/g, "") + " Documents";
          var trmListSystemName = trmListFullName.substring(0, 50);
          var trmDesc = term.get_description();

          // Build only the accordion headers and content placeholders
          //htmlText += '<div id="' + trmGUID + '_ee" class="panel panel-default akn-hidden">';
          htmlText +=
            '<div id="' +
            trmGUID +
            '_ee" class="panel panel-default akn-hidden">';
          htmlText +=
            '<div class="panel-heading accordion-toggle collapsed" aria-expanded="false" data-toggle="collapse" data-target="#collapse' +
            trmGUID +
            '" data-parent="#accordion">';
          htmlText += "<h3>" + trmName;
          htmlText +=
            ' <a class="spca" href="' +
            getAppRelativeUrl() +
            "/" +
            trmListSystemName +
            '" target="_blank" data-original-title="Click here to view resources ' +
            trmName +
            ' document library" data-toggle="tooltip">';
          htmlText +=
            '<img class="ms-rtePosition-3" src="' +
            scURL +
            '/PublishingImages/Inspect_2.svg" alt="" style="width: 20px; height: 20px;"/></a></h3>';
          htmlText +=
            '<div class="akn-panel-header-text">' + trmDesc + "</div>";
          htmlText += "</div>";
          //htmlText += '<div class="panel-collapse collapse" id="collapse' + trmGUID + '" aria-expanded="false" style="height: 0px;">';
          htmlText +=
            '<div class="panel-collapse collapse" id="collapse' +
            trmGUID +
            '" aria-expanded="false">';
          htmlText += '<div class="panel-body">';
          htmlText += '<div id="' + trmGUID + 'Root"></div>';
          htmlText += '<div class="panel-group">';
          htmlText += '<div id="' + trmGUID + '"></div>';
          htmlText += "</div>";
          htmlText += "</div>";
          htmlText += "</div>";
          htmlText += "</div>";
        }

        // Add accordion topic headers
        $("#accordion").append(htmlText);

        // Get all documents and links for content accordions
        for (var i = 0; i < terms.get_count(); i++) {
          var term = terms.getItemAtIndex(i);
          var trmGUID = term.get_id();
          var trmName = term.get_name();

          //getContentItems(trmGUID, trmName, sort);
          //Load Topic level items - Moved the getContentItems() call into separate function to be able to load sort order and protect async calls
          log("Time to load topic items for: " + trmName + ":" + trmGUID);
          loadTopicLevelLinksDocs(trmGUID, trmName);

          getSubtopicData(trmName, trmGUID);
        }

        // Get all site content links added to accordion
        //getLinks()
      },

      function(fArgs) {
        //fArgs[0] == fail callback sender
        //fArgs[1] == fail callback args.
        //in JSOM the callback args aren't used much -
        //the only useful one is probably the get_message()
        //on the fail callback
        var failmessage = fArgs[1].get_message();
        log(termName + " topic term not loaded. Error: " + failmessage);
      }
    );
  }

  // Get site GUID from script editor web page and run getTopicData
  var query =
    _spPageContextInfo.webAbsoluteUrl +
    "/_api/web/lists/GetByTitle('Pages')/items(" +
    _spPageContextInfo.pageItemId +
    ")";
  var siteGuid = "";
  var siteTitle = "";

  // Get Site GUID
  $.ajax({
    url: query,
    type: "GET",
    headers: {
      accept: "application/json;odata=verbose"
    },
    success: function(data) {
      siteTitle = _spPageContextInfo.webTitle; //data.d.Title;
      siteGuid = data.d.AKNSite.TermGuid;
      siteFullTitle = data.d.Title;
      if (siteGuid) {
        //$('#siteID').html("Page Title: " + siteTitle + " Site GUID: " + siteGuid);
        $("#accordion-container h2").html(siteFullTitle + " Resources");
        contacts();
        quicklinks();
        getTopicData(siteTitle, siteGuid);
        relatedlinks();
        //newsArticles();
      }
    },
    error: function(error) {
      alert("Error occurred");
    }
  });
}
