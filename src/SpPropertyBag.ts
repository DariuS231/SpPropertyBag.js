/*
 * SpPropertyBag.js
 * by rlv-dan (https://github.com/rlv-dan)
 * License: GPL3
*/
/// <reference path="../typings/sharepoint/SharePoint.d.ts" />
/// <reference path="../typings/microsoft-ajax/microsoft.ajax.d.ts" />

interface window {
    SpPropAdmin: any;
}


class SpPropertyBag{
    ctx: SP.ClientContext;
    web:any;
    allProperties:any;
    reloadRequired:boolean;
    constructor(){
        this.reloadRequired = false;
        
        this.ctx = SP.ClientContext.get_current();
        this.web = this.ctx.get_web();
        this.allProperties = this.web.get_allProperties();
        this.ctx.load(this.web);
        this.ctx.load(this.allProperties);
        
        
        let onSuccess:Function = Function.createDelegate(this,function(sender:any, err:any){this. showPropertiesDialog(this.allProperties.get_fieldValues());});
        let onError:Function = Function.createDelegate(this,function(sender:any, err:any){SP.UI.Notify.addNotification("Failed to get web properties...<br>" + err.get_message(), false);});
        this.ctx.executeQueryAsync(onSuccess, onError);
    }
    private executeChanges() {
		this.ctx.get_web().update();
		this.ctx.executeQueryAsync(function () {
			console.log("Web properties successfully modified");
		}, function () {
			console.error("Failed to set web property!");
		});
	};
	private setProperty(key:string, inputId:string) {
		let value = (<HTMLInputElement>document.getElementById(inputId)).value;
		this.allProperties.set_item(key, value);
		this.executeChanges();
	};
	private deleteProperty(key:string, inputId:string) {
		if (confirm('Are you sure you want to remove this property?')) {
			let table = document.getElementById(inputId).parentNode.parentNode;
			table.parentNode.removeChild(table);

			this.allProperties.set_item(key);
			this.executeChanges();
			this.reloadRequired = true;
		}
	};
	private addProperty() {
		let key = (<HTMLInputElement>document.getElementById("newKey")).value;
		let value = (<HTMLInputElement>document.getElementById("newValue")).value;
		(<HTMLInputElement>document.getElementById("newValue")).value = "";
		(<HTMLInputElement>document.getElementById("newKey")).value = "";
		this.allProperties.set_item(key, value);
		this.executeChanges();
	};

	private showPropertiesDialog(props: any) {
		let p:any;
        let type:string;
        let items:Array<any> = [];
		for(p in props) {
			if (props.hasOwnProperty(p)) {
				type = typeof(props[p]);
				if(type === "string") {
					items.push({"prop": p, "value": props[p].replace(/"/g, '&quot;')});
				}
			}
		}
		items.sort(function(a, b) {
			return a.prop.localeCompare(b.prop);
		});


		let html:HTMLElement = document.createElement('div');
		let h:string = '<hr><table style="margin: 1em;">';

		for(let i:number=0, itemsCount:number = items.length; i<itemsCount; i++) {
			h += '<tr>';
			h += '<td style="text-align: right; padding-top: 15px;"><b>' + items[i].prop + '</b></td>';
			h += '<td style="padding-top: 15px;"><input id="prop' + i + '" style="width:240px; " type="text" value="' + items[i].value + '"></inpu></td>';
			h += '<td style="padding-top: 15px;"><button onclick="SpPropAdmin.setProperty(\'' + items[i].prop + '\',\'prop' + i +'\'); return false;">Update</button></td>';
			h += '<td style="padding-top: 15px;"><button style="color: red; min-width: 1em;" onclick="SpPropAdmin.deleteProperty(\'' + items[i].prop + '\',\'prop' + i +'\'); return false;">X</button></td>';
			h += '</tr>';
		}
		h += '</table>';

		h += '<hr><h3>Add a new property:</h3>';
		h += '<div style="margin: 1em; padding-bottom: 2em;">Key: <input id="newKey"></inpu>';
		h += '&nbsp;&nbsp;&nbsp;Value: <input id="newValue"></inpu>';
		h += '&nbsp;<button onclick="SpPropAdmin.addProperty(); return false;">Add</button></div>';
		h += '<div></div>';

		html.innerHTML = h;

		SP.UI.ModalDialog.showModalDialog({
		 title: "Property Bag Editor",
		 html:html,
		 showClose: true,
		 autoSize: true,
		 dialogReturnValueCallback: function(dialogResult) {
			if(this.reloadRequired){
				window.location.reload();
			}
		 }
		});
	}

	
}

declare var SpPropAdmin:any;

SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function(){
    SpPropAdmin = new SpPropertyBag();
});
