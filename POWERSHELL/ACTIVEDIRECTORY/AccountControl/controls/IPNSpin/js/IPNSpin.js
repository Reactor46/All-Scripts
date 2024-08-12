var continueSpin = false,
	currentTarget = {
		field: null,
		spinner: null
	}
/**
Constructor for number spinner object

@param targetField - The target field object.
@param btnUp - The associated 'Up' button.
@param btnDn - The associated 'Down' button.
@param minValue - The minimum value.
@param maxValue - The maximum value.
@param oneToOne - Specifies if Up/Down buttons are tied-up to only one targetField or not. (boolean)
@param btnContainerId - The id container (span) where the Up/Down buttons are located.
@param interval - The interval between succeeding numbers. Optional. (default is 1)
@param hasLeadingZero - Specifies if single digit numbers must have leading zero or not. Optional. (default is false)
@param defaultValue - The default value. If entered value is not a number, default value will also be used. Optional. (default is the minValue)
@param decimalPlaces - The number of decimal places if interval is less than 1 (i.e. 0.5). (default is 0)
**/
function NumberSpinner (targetField, btnUp, btnDn, minValue, maxValue, oneToOne, btnContainerId, interval, hasLeadingZero, defaultValue, decimalPlaces) {
	var me = this;
	this.targetField = targetField;
	this.minValue = minValue;
	this.maxValue = maxValue;
	this.oneToOne = oneToOne;
	this.btnId = btnContainerId;
	this.interval = (typeof(interval) != 'undefined' && interval != null && interval > 0) ? interval : 1; //default is 1
	this.hasLeadingZero = (typeof(hasLeadingZero) != 'undefined' && hasLeadingZero != null) ? hasLeadingZero : false;	//default is false
	this.defaultValue = (typeof(defaultValue) == 'number') ? defaultValue : minValue;	//default is minValue
	this.decimalPlaces = (this.interval < 1 && decimalPlaces > 0) ? decimalPlaces : 0;	//default is 0

	//set default value
	targetField.value = ((me.hasLeadingZero && targetField.value < 10) ? "0" : "") + Number(this.defaultValue).toFixed(this.decimalPlaces);

	//attach event handlers automatically
	targetField.onclick = function () {
		setTarget(me.targetField,me)
	};
	targetField.onfocus = function () {
		setTarget(me.targetField,me)
	};
	targetField.onblur = function () {
		if (isNaN(parseInt(this.value,10))) {
			this.value = me.defaultValue;
		} else if (parseFloat(this.value) > me.maxValue) {
			this.value = me.maxValue;
		} else if (parseFloat(this.value) < me.minValue) {
			this.value = me.minValue;
		}
		this.value = ((me.hasLeadingZero && this.value < 10) ? "0" : "") + Number(this.value).toFixed(me.decimalPlaces); //>
		fakeBlur(currentTarget.field);
	}

    targetField.onkeydown = function (evt) {
        if ((evt ? evt.which : window.event.keyCode) == 38) {
            startSpin(1,me.btnId);
            return false;
        } else if ((evt ? evt.which : window.event.keyCode) == 40) {
            startSpin(0,me.btnId);
            return false;
        } else {
            return true;
        }
    };
    targetField.onkeyup = stopSpin;

if (!btnUp.onclick) {
		btnUp.onclick = function () {
			return false;
		}
	}
	if (!btnUp.onmousedown) {
		btnUp.onmousedown = function () {
			//targetField.focus();
			if (me.oneToOne) {
				setTarget(me.targetField,me);
			}
			startSpin(1,me.btnId);
		}
	}
	if (!btnUp.onmouseup) {
		btnUp.onmouseup = stopSpin;
	}
	if (!btnUp.onkeydown) {
		btnUp.onkeydown = function (evt) {
			if ((evt ? evt.which : window.event.keyCode) != 32) {
				return;
			}
			if (me.oneToOne) {
				setTarget(me.targetField,me);
			}
			startSpin(1,me.btnId);
		}
	}
	if (!btnUp.onkeyup) {
		btnUp.onkeyup = stopSpin;
	}
	if (!btnDn.onclick) {
		btnDn.onclick = function () {
			return false;
		}
	}
	if (!btnDn.onmousedown) {
		btnDn.onmousedown = function () {
			//targetField.focus();
			if (me.oneToOne) {
				setTarget(me.targetField,me);
			}
			startSpin(0,me.btnId);
		}
	}
	if (!btnDn.onmouseup) {
		btnDn.onmouseup = stopSpin;
	}
	if (!btnDn.onkeydown) {
		btnDn.onkeydown = function (evt) {
			if ((evt ? evt.which : window.event.keyCode) != 32) {
				return;
			}
			if (me.oneToOne) {
				setTarget(me.targetField,me);
			}
			startSpin(0,me.btnId);
		}
	}
	if (!btnDn.onkeyup) {
		btnDn.onkeyup = stopSpin;
	}
}

function setTime (mode, btnId) { //mode:1=up, 0=down
	//stop spinning
	if (!continueSpin || !currentTarget.field || (currentTarget.spinner && currentTarget.spinner.btnId != btnId)) {
		return;
	}
	//get current value
	var spinValue = parseFloat(currentTarget.field.value);

	//set default value if not numeric
	if (isNaN(spinValue)) {
		spinValue = currentTarget.spinner.minValue - currentTarget.spinner.interval;
	}
	//get next value
	spinValue = (mode == 0) ? spinValue - currentTarget.spinner.interval : spinValue + currentTarget.spinner.interval;

	//out of range?
	if (spinValue > currentTarget.spinner.maxValue) {
		spinValue = currentTarget.spinner.minValue;
	} else if (spinValue < currentTarget.spinner.minValue) {
		spinValue = currentTarget.spinner.maxValue;
	}
	//set decimal place
	spinValue = spinValue.toFixed(currentTarget.spinner.decimalPlaces);

	//set leading zero
	if (currentTarget.spinner.hasLeadingZero && spinValue < 10) { //>
		spinValue = "0" + spinValue;
	}
	//set value to target field
	currentTarget.field.value = spinValue;
	fakeFocus(currentTarget.field);

	//continue spinning
	setTimeout(function () {setTime(mode,btnId)},200);
}

function startSpin (mode, btnId) {
	continueSpin = true;
	setTime(mode, btnId);
}

function stopSpin () {
	continueSpin = false;
	if (currentTarget.field) {
		fakeFocus(currentTarget.field);
	}
}

function setTarget (focusedObj, spinnerObj) {
	fakeBlur(currentTarget.field);
	currentTarget.field = focusedObj;
	currentTarget.spinner = spinnerObj;
	if (focusedObj) {
		fakeFocus(focusedObj);
	}
}

function fakeFocus (el) {
	if (el != null) {
		el.className += ' focusedSpinnerInput';
	}
}

function fakeBlur (el) {
	if (el != null) {
		el.className = el.className.replace(/ focusedSpinnerInput/g, '');
	}
}

function resetTarget (e) {
	var targ = (e) ? e.target : event.srcElement;
	if (currentTarget.field != null && targ && targ.type != 'button' && (!targ.onmousedown || (targ.onmousedown && targ.onmousedown.toString().indexOf('setTarget') == -1))) {
		setTarget(null,null);
	}
}

//NOTE: if you have other document.onmousedown (or '<body onmousedown' in HTML) handler anywhere, you must combine them
document.onmousedown = resetTarget;

//fix for those browsers that don't support Number.toFixed()
//by adios of codingforums.com (http://www.codingforums.com/showthread.php?s=&threadid=29102#post149992)
if (!Number.prototype.toFixed) {
	Number.prototype.toFixed = function (decimals) {
		var decDigits = (isNaN(decimals)) ? 2 : decimals;
		var k = Math.pow(10, decDigits);
		var fixedNum = Math.round(parseFloat(this) * k) / k;
		var sFixedNum = new String(fixedNum);
		var aFixedNum = sFixedNum.split('.');
		var i = (aFixedNum[1]) ? aFixedNum[1].length : 0;
		if (i == 0 && decDigits) {
			sFixedNum += '.';
		}
		while (i++ < decDigits) {
			sFixedNum += '0';
		}
		return sFixedNum;
	}
}
/*************
END OF SCRIPT
**************/

/*********************
SAMPLE IMPLEMENTATION
var spinnerHour;
var spinnerMin;
var spinnerDay;

function initSpinner(){
	var f = document.timerForm;
	spinnerHour = new NumberSpinner(f.hhFrom,f.btnUp1,f.btnDn1,0,23,false,'btnTime1',1,true);
	spinnerMin = new NumberSpinner(f.mmFrom,f.btnUp1,f.btnDn1,0,59,false,'btnTime1',1,true);
	spinnerDay = new NumberSpinner(f.days,f.btnUp2,f.btnDn2,0,50,true,'btnTime2',0.5,true,5,2);
}

//NOTE: if you have other window.onload (or '<body onload' in HTML) handler anywhere, you must combine them
window.onload = initSpinner;
*********************/
