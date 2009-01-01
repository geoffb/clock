Private Function CancelEvent()

	With window.event
		.cancelBubble = True
		.returnValue = False
		.keyCode = 0
	End With

End Function