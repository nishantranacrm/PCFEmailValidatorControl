import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class PCFEmailValidatorControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {


	private _labelElement : HTMLLabelElement;
	private _container : HTMLDivElement;
	private _breakElement : HTMLBRElement;
	private _context : ComponentFramework.Context<IInputs>;
	private _notifyOutputChanged: () => void;

	private _textElement : HTMLInputElement;
	private _textElementChanged: EventListenerOrEventListenerObject;

	private _value: string;
	private _result : string;
	private _apiKey : string;

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
		this._notifyOutputChanged = notifyOutputChanged; 
		this._textElementChanged = this.emailAddressChanged.bind(this);



		this._value = "";
		
		if(context.parameters.emailProperty == null){
			this._value = "";
		}
		else
		{
			this._value = context.parameters.emailProperty.raw == null ? "" : context.parameters.emailProperty.raw;
		}
		
		if(context.parameters.apiProperty == null){
			this._apiKey = "";
		}
		else
		{
			this._apiKey = context.parameters.apiProperty.raw == null ? "" : context.parameters.apiProperty.raw;
		}

		this._textElement = document.createElement("input");
		this._textElement.setAttribute("type","text");		
		this._textElement.addEventListener("change", this._textElementChanged);	
		this._textElement.setAttribute("value",this._value);
		this._textElement.setAttribute("class", "InputText");			
		this._textElement.value = this._value;

		this._textElement.addEventListener("focusin", () => {
			this._textElement.className = "InputTextFocused";
			});
			this._textElement.addEventListener("focusout", () => {
				this._textElement.className = "InputText";
			});	


		this._breakElement = document.createElement("br");

		this._labelElement = document.createElement("label");		
		this._labelElement.setAttribute("border","2");	 
		
		this.CheckEmailValidity();
		
	
	 this._container.appendChild(this._textElement);
	 this._container.appendChild(this._breakElement);	 
	 this._container.appendChild(this._labelElement);

	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		let readOnly = this._context.mode.isControlDisabled; 
		let masked = false;	
	
		if (this._context.parameters.emailProperty.security) {
			readOnly = readOnly || !this._context.parameters.emailProperty.security.editable;
			masked = !this._context.parameters.emailProperty.security.readable;
		}
		if (masked)
			this._textElement.setAttribute("placeholder", "*******");
		else
			this._textElement.setAttribute("placeholder", "Insert an email address..");
		if (readOnly)
			this._textElement.readOnly = true;
		else
			this._textElement.readOnly = false;
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {
		    emailProperty : this._value
		};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
		this._textElement.removeEventListener("change", this._textElementChanged);	
	}

	public emailAddressChanged(evt: Event):void
	{
		this._value = this._textElement.value;
		this.CheckEmailValidity();
	
		this._notifyOutputChanged();		
	}

	private CheckEmailValidity() : void {

		this._labelElement.innerText = "loading..";


		var data = null;
		var result = null;

		if(this._textElement.value.length > 0 )
		{
			var xmlhttp = new XMLHttpRequest();
			xmlhttp.withCredentials = true;

			var _this = this;
			xmlhttp.addEventListener("readystatechange", function()
			{
				if(this.readyState === this.DONE)
				{

					var obj = JSON.parse(this.responseText);
					if (obj != null && obj.isValid != null) {

						if (obj.isValid === true)
							{
						_this._labelElement.innerText = "The email is valid !"						
						_this._labelElement.style.color = "white";
						_this._labelElement.style.backgroundColor= "#4CAF50";

						}
						else {
			
							_this._labelElement.innerText = "The email is invalid !"
							_this._labelElement.style.color = "white";
							_this._labelElement.style.backgroundColor = "#f44336";
						}			
					}			
					else {
						_this._labelElement.innerText = 'Error occured. Please try again later !';
						_this._labelElement.style.color = "black";
						_this._labelElement.style.backgroundColor = "#e7e7e7";
					}
				}				
			}
			)

			xmlhttp.open('GET', 'https://pozzad-email-validator.p.rapidapi.com/emailvalidator/validateEmail/'+  encodeURIComponent(this._textElement.value));
			xmlhttp.setRequestHeader("x-rapidapi-host", "pozzad-email-validator.p.rapidapi.com");
			xmlhttp.setRequestHeader("x-rapidapi-key", this._apiKey);
			xmlhttp.send(data);

			
		}
	}

}