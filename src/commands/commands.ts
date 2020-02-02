Office.onReady(() => {
	localStorage.clear();
});

function ValidateEmail(event) {
	Office.onReady(() => {
		Office.context.mailbox.getUserIdentityTokenAsync(function(asyncResult) {
			let token = asyncResult.value;
			localStorage.setItem('MicrosoftToken', token);
		});

		var mailboxItem = Office.context.mailbox.item;
		Office.context.mailbox.item.saveAsync(function(result) {
			localStorage.setItem('EWSURL', Office.context.mailbox.ewsUrl);
			localStorage.setItem('ITEMID', result.value);
			localStorage.setItem('ItemType', mailboxItem.itemType);
			localStorage.setItem('IsAuthTokenRequest', 'N');

			if (mailboxItem.itemType === Office.MailboxEnums.ItemType.Appointment) {
				mailboxItem.organizer.getAsync({ asyncContext: event }, function(asyncResult) {
					localStorage.setItem('FROM', JSON.parse(JSON.stringify(asyncResult.value))['emailAddress']);
				});
			}

			Office.context.mailbox.getCallbackTokenAsync(function(asyncResult) {
				let token = asyncResult.value;
				localStorage.setItem('EmailItemToken', token);
				OpenDialog(event);
			});
		});
	});
}

function OpenDialog(event) {
	Office.context.ui.displayDialogAsync(
		window.location.origin + '/taskpane.html',
		{ height: 70, width: 50, displayInIframe: true },
		function(result) {
			var dialog = result.value;
			console.log(result);
			if (dialog !== undefined) {
				dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(args) {
					dialog.close();
					var msg = JSON.parse(args.message);
					if (msg.message === 'send' && event.source.id !== 'msgReadOpenPaneButton') {
						event.completed({ allowEvent: true });
					} else {
						event.completed({ allowEvent: false });
					}
				});
				dialog.addEventHandler(Office.EventType.DialogEventReceived, function handleAuthDialogMessage(args) {
					if (args.error !== undefined && args.error > 0) {
						dialog.close();
						event.completed({ allowEvent: false });
					}
				});
			}
		}
	);
}

function getGlobal() {
	return typeof self !== 'undefined'
		? self
		: typeof window !== 'undefined' ? window : typeof global !== 'undefined' ? global : undefined;
}
const g = getGlobal() as any;
g.ValidateEmail = ValidateEmail;
