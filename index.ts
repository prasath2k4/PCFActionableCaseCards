import {IInputs, IOutputs} from "./generated/ManifestTypes";
import * as $ from 'jquery';
import Swal from 'sweetalert2';


import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;

export class CaseCards implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	
	private _container: HTMLDivElement;
	private _context: ComponentFramework.Context<IInputs>;
	public _resolveCase: EventListenerOrEventListenerObject;
	public _openCase: EventListenerOrEventListenerObject;
	
	private _addNotesButton: HTMLElement;
	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		// Add control initialization code
		this._context = context;
		this._container = container;
	
		
	}	

	public OpenCase(evt: Event): void {
		//@ts-ignore
		var incID = evt.target.id;

		//@ts-ignore
		window.open(this._context.page.getClientUrl()+'/main.aspx?appid='+this._context.page.appId+'&pagetype=entityrecord&etn=incident&id='+incID,'_blank');				
	}

	public ResolveCase(evt: Event): void {
		//@ts-ignore
		var incID = evt.target.id;

		var resolutionRemarks;

		Swal.mixin({
			input: 'text',
			confirmButtonText: 'Proceed to resolve &rarr;',
			showCancelButton: true,
			progressSteps: ['1', '2', '3']
		  }).queue([
			{
				title: 'Case Resolution',
				text: 'Please enter the resolution remarks'
			  },
			  {
				title: 'Billable Time',
				text: 'Please enter the billable time in minutes'
			  },
			  {
				title: 'Customer Experience',
				text: 'How was the experience interacting with this customer?'
			  } 
		  ]).then((result) => {
			if (result.value) {

				var answers = JSON.stringify(result.value);
				var answersArray = answers.substring(1, answers.length-1).split(",");


				var incidentresolution = {
					"subject": "Put Your Resolve Subject Here",
					"incidentid@odata.bind": "/incidents("+incID+")",//Id of incident
					"timespent": Number(answersArray[1].substring(1, answersArray[1].length-1)),//This is billable time in minutes
					"description": answersArray[0]
				};
				 
				var parameters = {
					"IncidentResolution": incidentresolution,
					"Status": -1
				};
				 
							 
				var req = new XMLHttpRequest();
				
				//@ts-ignore
				req.open("POST", this._context.page.getClientUrl() + "/api/data/v8.2/CloseIncident", true);
			
				req.setRequestHeader("OData-MaxVersion", "4.0");
				req.setRequestHeader("OData-Version", "4.0");
				req.setRequestHeader("Accept", "application/json");
				req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
				req.onreadystatechange = function () {
					if (this.readyState === 4) {
						req.onreadystatechange = null;
						if (this.status === 204) {
							//Success - No Return Data - Do Something
						} else {
							var errorText = this.responseText;
							//Error and errorText variable contains an error - do something with it
						}
					}
				};
				req.send(JSON.stringify(parameters));
				$('#card'+incID).hide();


			  Swal.fire({
				title: 'All done!',
				html: 'The case is resolved',
				confirmButtonText: 'OK!'
			  });
			}
		  });	

	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		$('.wrapper').empty();

		let resolveCase = this.ResolveCase.bind(this);

		let openCaseForm = this.OpenCase.bind(this);

		let wrapperDiv = document.createElement("div");
		wrapperDiv.className = "wrapper";

		if(context.parameters.sampleDataSet.sortedRecordIds.length == 0){
		this._container.innerHTML = "No more cases! Go have fun!";
		this._container.setAttribute("style","font-size:-webkit-xxx-large;text-align:center;padding:180px");
		}


		$.each( context.parameters.sampleDataSet.sortedRecordIds, function( index, value ){

			let cardDiv = document.createElement("div");
			cardDiv.className = "card";
			cardDiv.id = "card"+value.toString();

			let checkboxInput = document.createElement("input");
			checkboxInput.type = "checkbox";
			checkboxInput.id = "card"+index.toString();
			checkboxInput.className = "more";

			let contentDiv = document.createElement("div");
			contentDiv.className = "content";

			let frontDiv = document.createElement("div");
			frontDiv.className = "front";

			let frontDivInner = document.createElement("div");
			frontDivInner.className = "inner";

			let caseTitleDiv = document.createElement("H3");
			caseTitleDiv.setAttribute("style","color:black");
			caseTitleDiv.innerText = context.parameters.sampleDataSet.records[value].getFormattedValue("title");

			let caseNumDiv = document.createElement("H2");
			caseNumDiv.setAttribute("style","color:black;padding:30px	");
			caseNumDiv.innerText = context.parameters.sampleDataSet.records[value].getFormattedValue("ticketnumber");

			let frontDivInnerH3 = document.createElement("H3");

			let caseCreatedOnDiv = document.createElement("H4");
			caseCreatedOnDiv.innerText = "Received on :"+context.parameters.sampleDataSet.records[value].getFormattedValue("createdon");

			let caseSeverityDiv = document.createElement("H4");
			caseSeverityDiv.setAttribute("style","padding:38px");
			caseSeverityDiv.innerText = "Severity : "+context.parameters.sampleDataSet.records[value].getFormattedValue("prioritycode");

			let detailsLbl = document.createElement("label");
			detailsLbl.setAttribute("for","card"+index.toString());
			detailsLbl.className = "button";
			detailsLbl.innerText = "Details";

			let openCase = document.createElement("input");
			openCase.type = "button";
			openCase.className = "button";
			openCase.setAttribute("value","Open Case");
			openCase.id = value;
			openCase.addEventListener("click",openCaseForm);

			frontDivInner.appendChild(caseNumDiv);
			frontDivInner.appendChild(caseTitleDiv);
			frontDivInner.appendChild(frontDivInnerH3);
			frontDivInner.appendChild(caseCreatedOnDiv);
			frontDivInner.appendChild(caseSeverityDiv);
			frontDivInner.appendChild(detailsLbl);
			
			frontDiv.appendChild(frontDivInner);
			frontDiv.appendChild(openCase);

			let backDiv = document.createElement("div");
			backDiv.className = "back";

			let backDivInner = document.createElement("div");
			backDivInner.className = "inner";
			backDivInner.setAttribute("style","font-size:initial");

			let descriptionDiv = document.createElement("div");
			descriptionDiv.className = "description";

			let descriptionPTag = document.createElement("p");
			descriptionPTag.innerText = "Description";

			let descriptionPTagMore = document.createElement("p");
			descriptionPTagMore.innerText = context.parameters.sampleDataSet.records[value].getFormattedValue("description");

			descriptionDiv.appendChild(descriptionPTag);
			descriptionDiv.appendChild(descriptionPTagMore);

			let caseSeverityDivBack = document.createElement("div");
			caseSeverityDivBack.className = "location";
			caseSeverityDivBack.innerText = "Severity : "+context.parameters.sampleDataSet.records[value].getFormattedValue("prioritycode");

			let DueOnDiv = document.createElement("div");
			DueOnDiv.className = "price";
			DueOnDiv.innerText = "Due : "+context.parameters.sampleDataSet.records[value].getFormattedValue("createdon");

			//<label for="card1" class="button return" aria-hidden="true">
			let returnLbl = document.createElement("label");
			returnLbl.setAttribute("for","card"+index.toString());
			returnLbl.className = "button return";

			let arrowLeft = document.createElement("i");
			arrowLeft.className = "fas fa-arrow-left";

			returnLbl.appendChild(arrowLeft);

			backDivInner.appendChild(descriptionDiv);
			backDivInner.appendChild(caseSeverityDivBack);
			backDivInner.appendChild(DueOnDiv);
			backDivInner.appendChild(returnLbl);

			let boxDiv = document.createElement("div");
			boxDiv.className = "box";	
			
			let resolveBtn = document.createElement("input");
			resolveBtn.id = value;
			resolveBtn.type = "button";
			resolveBtn.className = "button";
			resolveBtn.innerText = "Resolve the case";
			resolveBtn.setAttribute("value","Resolve the case");
			resolveBtn.addEventListener("click",resolveCase);
						
			boxDiv.appendChild(resolveBtn);	

			backDiv.appendChild(backDivInner);
			backDiv.appendChild(boxDiv);

			contentDiv.appendChild(frontDiv);
			contentDiv.appendChild(backDiv);

			cardDiv.appendChild(checkboxInput);
			cardDiv.appendChild(contentDiv);

			wrapperDiv.appendChild(cardDiv);
		});

		this._container.appendChild(wrapperDiv);

	}
	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
		
		
	}

}