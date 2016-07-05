var hpsRoot = hpsRoot || {};

hpsRoot.CustomActionUtility = {

	getAllCustomActions : function () {

		var deferred = $.Deferred();

		hpsRoot.CustomActionUtility.getWebCustomActions().done(function (wa) {

			hpsRoot.CustomActionUtility.getSiteCustomActions().done(function (sa) {

				var array = wa.concat(sa).sort();

				deferred.resolve(array);

			});

		});

		return deferred;

	},

	getWebCustomActions : function () {
		return hpsRoot.CustomActionUtility.getCustomActions("web");
	},

	getSiteCustomActions : function () {
		return hpsRoot.CustomActionUtility.getCustomActions("site");
	},

	getCustomActions : function (siteOrWeb) {

		var deferred = $.Deferred();

		siteOrWeb = (siteOrWeb == "site" ? "site" : "web");

		$.get({
			url : _spPageContextInfo.webAbsoluteUrl + "/_api/" + siteOrWeb + "/userCustomActions?$orderby=Sequence",
			dataType : "json",
			contentType : 'application/json',
			headers : {
				"Accept" : "application/json; odata=verbose"
			},
			cache : false,
			success : function (response) {
				deferred.resolve(response.d.results);
			},
			error : function (response) {
				deferred.reject(response);
			}
		});

		return deferred;

	},

	installCustomAction(siteOrWeb, title, url, seq, absolute, id) {

		var deferred = $.Deferred();

		var webContext = SP.ClientContext.get_current();
		var userCustomActions = siteOrWeb == "site" ? webContext.get_site().get_userCustomActions() : webContext.get_web().get_userCustomActions();

		webContext.load(userCustomActions);
		seq = parseInt(seq) || 1000;
		var action = userCustomActions.add();

		action.set_location("ScriptLink");
		action.set_title(title);
		var isCss = /\.css$/gi.test(url);
		var block;
		if (isCss) {
			block = [
				"(function(){",
				"var head1 = document.getElementsByTagName('head')[0];",
				"var link1 = document.createElement('link');",
				"link1.type = 'text/css';",
				"link1.rel = 'stylesheet';",
				"link1.href = '" + (absolute ? url : _spPageContextInfo.webServerRelativeUrl + "/" + url) + "';",
				(id ? "link1.id = '" + id + "';" : ""),
				"head1.appendChild(link1);",
				"})();"
			].join("");
			action.set_scriptBlock(block);
			action.set_description((absolute ? url : _spPageContextInfo.webServerRelativeUrl + "/" + url));
		} else if (id || absolute) {
			block = [
				"(function(){",
				"var head1 = document.getElementsByTagName('head')[0];",
				"var script1 = document.createElement('script');",
				"script1.type = 'text/javascript';",
				"script1.src = '/" + (absolute ? url : _spPageContextInfo.webServerRelativeUrl + "/" + url) + "';",
				"script1.id = '" + id + "';",
				"head1.appendChild(script1);",
				"})();"
			].join("");
			action.set_scriptBlock(block);
			action.set_description((absolute ? url : _spPageContextInfo.webServerRelativeUrl + "/" + url));
		} else {
			action.set_scriptSrc((absolute ? url : "~sitecollection/" + url));
			action.set_description((absolute ? url : "~sitecollection/" + url));
		}
		action.set_sequence(seq);
		action.update();

		webContext.load(action);

		webContext.executeQueryAsync(function (response) {
			deferred.resolve(response);
		}, function (response) {
			deferred.reject(response);
		});

		return deferred;

	},

	uninstallCustomAction : function (siteOrWeb, id, url) {
		var deferred = $.Deferred();
		var webContext = SP.ClientContext.get_current();
		var userCustomActions = siteOrWeb == "site" ? webContext.get_site().get_userCustomActions() : webContext.get_web().get_userCustomActions();
		webContext.load(userCustomActions);
		webContext.executeQueryAsync(function () {
			var count = userCustomActions.get_count();
			for (var i = count - 1; i >= 0; i--) {
				var action = userCustomActions.get_item(i);
				if (String(id).length > 0) {
					if (action.get_id().toString() == id) {
						action.deleteObject();
					}
				} else {
					if (action.get_scriptSrc() == "~sitecollection/" + url || action.get_title() == url) {
						action.deleteObject();
					}
				}
			}
			webContext.executeQueryAsync(function (response) {
				deferred.resolve(response);
			}, function (response) {
				deferred.reject(response);
			});
		}, function (arguments) {
			deferred.reject(response);
		});
		return deferred;
	},

	init : function () {

		SP.SOD.executeFunc("//ajax.aspnetcdn.com/ajax/jQuery/jquery-3.0.0.min.js", "jQuery", init);

		function init() {

			function loadCustomActionsIntoTable() {

				hpsRoot.CustomActionUtility.getAllCustomActions().done(function (actions) {

					$("#hpsCustomActionTable").find("tbody").empty();

					actions.forEach(function (action) {

						var scopeText = (action.Scope == 3 ? "Web" : "Site");

						var $tr = $('<tr><td>' + action.Title + '</td><td>' + action.Description + '</td><td>' + action.Sequence + '</td><td>' + scopeText + '</td><td><button class="hpsDeleteCustomActionButton" type="button">x</button></td></tr>');

						$("#hpsCustomActionTable").find("tbody").append($tr);

						$tr.find(".hpsDeleteCustomActionButton").click(function () {

							if (confirm("delete?")) {

								hpsRoot.CustomActionUtility.uninstallCustomAction(scopeText.toLowerCase(), action.Id).done(function () {

									loadCustomActionsIntoTable();

								});

							}

						});

					});

					if (actions.length == 0) {
						$("#hpsCustomActionTable").find("tbody").append('<tr><td colspan="5">No Custom Actions</td></tr>');
					}

				});

			}

			$.getScript(_spPageContextInfo.siteAbsoluteUrl + "/_layouts/15/sp.js").done(function () {

				$("#DeltaPlaceHolderMain").empty().load("//raw.githubusercontent.com/supergometan/CA_bookmarklet/master/form.html", function (data) {

					$("button#install-site-user-custom-action").click(function () {
						installUserCustomAction("site");
					});
					$("button#uninstall-site-user-custom-action").click(function () {
						uninstallUserCustomAction("site");
					});
					$("button#install-web-user-custom-action").click(function () {
						installUserCustomAction("web");
					});
					$("button#uninstall-web-user-custom-action").click(function () {
						uninstallUserCustomAction("web");
					});

					loadCustomActionsIntoTable();

					$("#hpsNewCustomActionButton").click(function () {

						SP.SOD.executeFunc("sp.ui.dialog.js", "SP.UI.ModalDialog.showModalDialog", function () {
							var modal = SP.UI.ModalDialog.showModalDialog({
									title : "Create Custom Action",
									html : $($("#hpsCustomActionModalContent").html())[0]
								});

							var $modal = $(modal.get_html());

							$modal.find("#hpsCustomActionModalCancelButton").click(function () {

								modal.close(modal);

							});

							$modal.find("#hpsCustomActionModalSaveButton").one("click", function () {

								var scope = "";
								$modal.find("input[name=scope]").each(function () {
									if ($(this).prop("checked"))
										scope = $(this).val();
								});
								
								var absolute = $modal.find("input[name=absolute]").prop("checked");

								var title = $modal.find("input[name=title]").val();
								var url = $modal.find("input[name=url]").val().trim();
								var sequence = parseInt($modal.find("input[name=sequence]").val());

								hpsRoot.CustomActionUtility.installCustomAction(scope, title, url, sequence, absolute).done(function () {

									modal.close(modal);

									loadCustomActionsIntoTable();

								});

							});
						});

					});

				});

			});

		}
	}

};
_spBodyOnLoadFunctions.push(hpsRoot.CustomActionUtility.init);

hpsRoot.CustomActionUtility.init();
