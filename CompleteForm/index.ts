import { time } from "console";
import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class CompleteForm implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _context: ComponentFramework.Context<IInputs>;

	private _container: HTMLDivElement;
	private _percentsDiv: HTMLDivElement;
	private _attributesDiv: HTMLDivElement;

	private _geralPercentLabel: HTMLLabelElement;

	private _noneNulls: Xrm.Attributes.Attribute[];
	private _noneNotNulls: Xrm.Attributes.Attribute[];
	private _noneDiv: HTMLDivElement;
	private _noneLabel: HTMLLabelElement;

	private _recommendedNulls: Xrm.Attributes.Attribute[];
	private _recommendedNotNulls: Xrm.Attributes.Attribute[];
	private _recommendedDiv: HTMLDivElement;
	private _recommendedLabel: HTMLLabelElement;

	private _requiredNulls: Xrm.Attributes.Attribute[];
	private _requiredNotNulls: Xrm.Attributes.Attribute[];
	private _requiredDiv: HTMLDivElement;
	private _requiredLabel: HTMLLabelElement;

	private _formContext: any;

	/**
	 * Empty constructor.
	 */
	constructor() {

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		this.GetFormContext();
		this.GetAttributes();

		//Render HTML
		this._container = document.createElement("div");
		container.append(this._container); //Main
		this.RenderHTML();

		this.AssociateValues();
	}

	private GetFormContext() {
		this._formContext = (window as any).getCurrentXrm();
	}

	private GetAttributes() {

		let attributes: Xrm.Attributes.Attribute<any>[] = this._formContext._page.getAttribute();

		this._noneNulls = attributes.filter(f => f.getRequiredLevel() == "none" && !f.getValue());
		this._noneNotNulls = attributes.filter(f => f.getRequiredLevel() == "none" && f.getValue());
		this._recommendedNulls = attributes.filter(f => f.getRequiredLevel() == "recommended" && !f.getValue());
		this._recommendedNotNulls = attributes.filter(f => f.getRequiredLevel() == "recommended" && f.getValue());
		this._requiredNulls = attributes.filter(f => f.getRequiredLevel() == "required" && !f.getValue());
		this._requiredNotNulls = attributes.filter(f => f.getRequiredLevel() == "required" && f.getValue());
	}

	private RenderHTML() {
		// %
		this._geralPercentLabel = document.createElement("label");
		this._container.append(this._geralPercentLabel);

		//Required Levels
		this._percentsDiv = document.createElement("div");
		this._container.append(this._percentsDiv);
		{
			//None
			this._noneDiv = document.createElement("div");
			this._noneDiv.setAttribute("class", "divNone");
			this._percentsDiv.append(this._noneDiv);
			{
				this._noneLabel = document.createElement("label");
				this._noneLabel.setAttribute("class", "labelDetail");
				//this._noneLabel.addEventListener("click", this.GetNoneNullAttributes.bind(this));
				this._noneDiv.append(this._noneLabel);
			}

			//Recommended
			this._recommendedDiv = document.createElement("div");
			this._recommendedDiv.setAttribute("class", "divRecommended");
			this._percentsDiv.append(this._recommendedDiv);
			{
				this._recommendedLabel = document.createElement("label");
				this._recommendedLabel.setAttribute("class", "labelDetail");
				//this._recommendedLabel.addEventListener("click", this.GetRecommendedNullAttributes.bind(this));
				this._recommendedDiv.append(this._recommendedLabel);
			}

			//Required
			this._requiredDiv = document.createElement("div");
			this._requiredDiv.setAttribute("class", "divRequired");
			this._percentsDiv.append(this._requiredDiv);
			{
				this._requiredLabel = document.createElement("label");
				this._requiredLabel.setAttribute("class", "labelDetail");
				//this._requiredLabel.addEventListener("click", this.GetRequiredNullAttributes.bind(this));
				this._requiredDiv.append(this._requiredLabel);
			}
		}

		//Break Row
		let br = document.createElement("br");
		this._container.append(br);
		this._container.append(br);

		//Attributes Div
		this._attributesDiv = document.createElement("div");
		this._attributesDiv.setAttribute("class", "divAttributes");
		this._container.append(this._attributesDiv);
	}

	private AssociateValues() {
		if (this.GetAllAttributes() == 0)
			return;

		//% Completed
		this._geralPercentLabel.innerText = ((1 - (this.GetNullAttributes() / this.GetAllAttributes())) * 100).toFixed(0).toString() + " % Completed";

		//x/y by Required Level
		//this._noneLabel.innerText = this._noneNotNulls.length + "/" + this.GetNones();
		this._noneDiv.title = (this.GetNones() - this._noneNotNulls.length) + this.GetAttributesLabel(this._noneNulls.map(m => m.getName()));
		//this._recommendedLabel.innerText = this._recommendedNotNulls.length + "/" + this.GetRecommendeds();
		this._recommendedDiv.title = (this.GetRecommendeds() - this._recommendedNotNulls.length) + this.GetAttributesLabel(this._recommendedNulls.map(m => m.getName()));
		//this._requiredLabel.innerText = this._requiredNotNulls.length + "/" + this.GetRequireds();
		this._requiredDiv.title = (this.GetRequireds() - this._requiredNotNulls.length) + this.GetAttributesLabel(this._requiredNulls.map(m => m.getName()));

		//% by Required Level
		this._noneDiv.style.width = ((this.GetNones() / this.GetAllAttributes()) * 100).toString().replace(",", ".") + "%";
		this._recommendedDiv.style.width = ((this.GetRecommendeds() / this.GetAllAttributes()) * 100).toString().replace(",", ".") + "%";
		this._requiredDiv.style.width = ((this.GetRequireds() / this.GetAllAttributes()) * 100).toString().replace(",", ".") + "%";
	}

	public GetAttributesLabel(attributes: Array<string>): string {
		let labels: Array<string> = new Array<any>();
		attributes.forEach(attribute_ => {
			var control = this._formContext._page.getControl(attribute_);
			if (control)
				labels.push(control.getLabel());
			else
				labels.push(attribute_);
		});
		return " null attribute(s) \r\n" + labels.join("\r\n");
	}

	public GetAllAttributes(): number {
		return this._requiredNulls.length + this._recommendedNulls.length + this._noneNulls.length + this._requiredNotNulls.length + this._recommendedNotNulls.length + this._noneNotNulls.length;
	}

	public GetNullAttributes(): number {
		return this._requiredNulls.length + this._recommendedNulls.length + this._noneNulls.length;
	}

	public GetNones(): number {
		return this._noneNulls.length + this._noneNotNulls.length;
	}

	public GetRecommendeds(): number {
		return this._recommendedNulls.length + this._recommendedNotNulls.length;
	}

	public GetRequireds(): number {
		return this._requiredNulls.length + this._requiredNotNulls.length;
	}

	private GetNoneNullAttributes() {
		this.RenderAttributes(this._noneNulls.map(m => m.getName()), "grey");
	}

	private GetRecommendedNullAttributes() {
		this.RenderAttributes(this._recommendedNulls.map(m => m.getName()), "dodgerblue");
	}

	private GetRequiredNullAttributes() {
		this.RenderAttributes(this._requiredNulls.map(m => m.getName()), "tomato");
	}

	private RenderAttributes(attributeLogicalNames: string[], backgroundColor: string) {
		//Clear Attribute List
		this.ClearAttributes();

		//Set Color
		this._attributesDiv.style.backgroundColor = backgroundColor;

		//Attribute List
		let ul: HTMLUListElement;
		ul = document.createElement("ul");
		this._attributesDiv.append(ul);

		let items: HTMLLIElement[];
		items = new Array();

		//Attribute Item
		for (let index = 0; index < attributeLogicalNames.length; index++) {
			const attributeItem = attributeLogicalNames[index];

			let li: HTMLLIElement;
			li = document.createElement("li");
			li.id = attributeItem;
			li.innerText = this._formContext._page.getControl(attributeItem) != null ? this._formContext._page.getControl(attributeItem).getLabel() : attributeItem;
			li.addEventListener("click", this.SetFocus.bind(this, li.id));
			items.push(li);
		}

		//Sort By DisplayName
		items = items.sort(function compare(a, b) {
			if (a.innerText < b.innerText) {
				return -1;
			}
			if (a.innerText > b.innerText) {
				return 1;
			}
			return 0;
		});

		//Insert Itens
		for (let index = 0; index < items.length; index++) {
			const li = items[index];
			ul.append(li);
		}
	}

	private SetFocus(attribute: string) {
		alert(attribute);
		this.ClearAttributes();
		//Xrm.Page.getControl(attribute)!.setFocus();
	}

	public ClearAttributes() {
		while (this._attributesDiv.firstChild) {
			this._attributesDiv.removeChild(this._attributesDiv.firstChild);
		}
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
	}
}