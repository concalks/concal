/* Central Logic */
export function initialize(): void {
    console.log('>>> taskpane.js bundle executed');

    /*
        IMPORTANT!
        DO NOT EDIT THE CONSTANTS BELOW, UNLESS YOU ADJUST ALL REGEX PATTERNS, HERE AND IN THE DESKTOP APP!
     */
    const START_AUTOMATED = "[START AUTOMATED SECTION: DO NOT EDIT MANUALLY BELOW THIS LINE!]";  // DO NOT CHANGE WITHOUT CAREFUL TESTING
    const START_HEADERS = "[START_CONTACT_HEADERS]";  // DO NOT CHANGE WITHOUT CAREFUL TESTING
    const END_HEADERS = "[END_CONTACT_HEADERS]";      // DO NOT CHANGE WITHOUT CAREFUL TESTING
    const START_DETAILS = "[START_CONTACT_DETAILS]";  // DO NOT CHANGE WITHOUT CAREFUL TESTING
    const END_DETAILS = "[END_CONTACT_DETAILS]";      // DO NOT CHANGE WITHOUT CAREFUL TESTING
    const FIXED_HEADERS = "Verarbeitungszweck;Ursprungsquelle;Letzter Austausch";  // DO NOT CHANGE
    const FIXED_DETAILS = "Kundenbetreuung;Concal;";                               // DO NOT CHANGE
    /*
        END OF CRITICAL CONSTANTS BLOCK
     */

    interface ContactFieldConfig {
        id: string;
        cobraId: string;
        label: string;
        placeholder: string;
    }

    const INPUT_FIELDS: ContactFieldConfig[] = [ //todo weitere Felder einfach hier hinzufügen, sie werden automatisch angezeigt!
        {id: "institution", cobraId: "Institution", label: "Institution", placeholder: "München e.V."},
        {id: "institution2", cobraId: "Institution2", label: "Weitere<br>Institution", placeholder: "Bayerischer Rat"},
        {id: "abteilung", cobraId: "Abteilung", label: "Abteilung", placeholder: "IT-Abteilung"},
        {id: "anrede", cobraId: "Anrede", label: "Anrede", placeholder: "Frau"},
        {id: "titel", cobraId: "AkademischerTitel", label: "Akademischer Titel", placeholder: "Dr."},
        {id: "vorname", cobraId: "Vorname", label: "Vorname", placeholder: "Erika"},
        {id: "nachname", cobraId: "Nachname", label: "Nachname", placeholder: "Mustermann"},
        {id: "position", cobraId: "Position", label: "Position", placeholder: "Generaldirektorin"},
        {id: "strasse", cobraId: "Straße", label: "Straße + Haus Nr.", placeholder: "Musterstraße 1"},
        {id: "postfach", cobraId: "Postfach", label: "Postfach", placeholder: "012345"},
        {id: "gebiet", cobraId: "Gebiet", label: "Gebiet", placeholder: ""},
        {id: "plz", cobraId: "PLZ", label: "PLZ", placeholder: "87654"},
        {id: "ort", cobraId: "Ort", label: "Ort", placeholder: "Musterstadt"},
        {id: "tel_zent", cobraId: "Telefon zentral", label: "Telefon zentral", placeholder: "08765 / 987654"},
        {id: "tel_dir", cobraId: "Telefon direkt", label: "Telefon direkt", placeholder: "08765 / 987653"},
        {id: "mobil", cobraId: "Mobil", label: "Mobil", placeholder: "+49 123 456789876"},
        {id: "fax", cobraId: "Telefax", label: "Fax", placeholder: "08765 / 987653"},
        {id: "web", cobraId: "Webseite", label: "Website", placeholder: "https://musterseite.de"},
        {id: "mail_pers", cobraId: "E-Mail persönlich", label: "E-Mail persönlich", placeholder: "erika@mustermann.de"},
        {id: "mail_zent", cobraId: "E-Mail zentral", label: "E-Mail zentral", placeholder: "erika.mustermann@work.de"},
        {id: "einladung", cobraId: "Einladung", label: "Einladung", placeholder: ""}
    ];

    Office.onReady((info) => {
        if (info.host === Office.HostType.Outlook) {
            const addContactB = document.getElementById("addContact");
            if (addContactB) {
                addContactB.addEventListener("click", addContact);
            } else {
                console.error("No addContact button found!")
            }
            const expContactB = document.getElementById("exportContact");
            if (expContactB) {
                expContactB.addEventListener("click", exportContactData);
            } else {
                console.error("No exportContact button found!")
            }

            Office.context.mailbox.item?.subject.getAsync(
                function (result: any) {
                    const heading = document.getElementById("heading");
                    if (heading && result.value) {
                        heading.textContent = `Kontaktinformationen zu ${sanitizeInput(result.value)}`;
                    }
                });
        }
    });

    function addContact() {
        console.log("Adding contact details");

        const addContactButton = document.getElementById("addContact") as HTMLButtonElement;
        if (addContactButton && addContactButton instanceof HTMLButtonElement) {
            addContactButton.classList.add("ms-Button--disabled");
            addContactButton.disabled = true;
        } else {
            console.error("Something went very wrong with addContact button!");
        }

        const expContactButton = document.getElementById("exportContact") as HTMLButtonElement;
        if (expContactButton && expContactButton instanceof HTMLButtonElement) {
            expContactButton.classList.add("ms-Button--disabled");
            expContactButton.disabled = true;
        } else {
            console.error("Something went very wrong with exportContact button!");
        }

        const dyn = document.getElementById("dyn"); // insert html fields in placeholder element
        const container = document.createElement("div");
        container.id = "container";
        let inputFields = "";
        INPUT_FIELDS.forEach(function (value) {
            inputFields += `
               <label class="ms-Label">${value.label}</label>
               <input class="ms-TextField-field" type="text" id="${value.id}" placeholder="${value.placeholder}" />
                `;
        });

        container.innerHTML = `
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px 8px; margin-bottom: 1em">
            ${inputFields}
        </div>
        <button role="button" id="save" class="ms-Button--primary">
            <span class="ms-Button-label">Speichern</span>
        </button>
        <button role="button" id="discard" class="ms-Button--primary">
            <span class="ms-Button-label">Abbrechen</span>
        </button>
      `;

        dyn?.appendChild(container);
        document.getElementById("save")?.addEventListener("click", saveContact);
        document.getElementById("discard")?.addEventListener("click", discard);

    }

    async function saveContact() {
        let contactHeaders = "";
        let contactDetails = "";
        INPUT_FIELDS.forEach(function (value) {
            contactHeaders += value.cobraId + ";"
            const rawValue = (document.getElementById(value.id) as HTMLInputElement)?.value || "";
            contactDetails += sanitizeInput(rawValue) + ";";
        })
        contactHeaders += FIXED_HEADERS;
        contactDetails += FIXED_DETAILS;

        const subject = await getAsyncPromise<string>((cb) =>
            Office.context.mailbox.item?.subject.getAsync(cb)
        ).catch(() => "");

        const startDate = await getAsyncPromise<Date>((cb) =>
            Office.context.mailbox.item?.start.getAsync(cb)
        ).catch(() => null);

        let lastMet = "";
        if (startDate) {
            lastMet = `Am ${startDate.toLocaleDateString()} zu `;
            lastMet += subject ? `dem Termin: ${sanitizeInput(subject)}` : "einem Termin";
        }

        const fullContactDetails = `${contactDetails}${lastMet};`;
        await putEventBody(contactHeaders, fullContactDetails);
        discard();
    }

    async function putEventBody(contactHeaders: string, contactDetails: string) {
        const currentData = await getAsyncPromise<string>((cb) =>
            Office.context.mailbox.item?.body.getAsync(Office.CoercionType.Html, cb)).catch(() => discard()) || "";

        let headers = "";
        if (!currentData.includes(START_HEADERS)) {
            headers = `<div style="color: #AAAAAA;"><br><br><br>---${START_AUTOMATED}---<br>---${START_HEADERS}---<br>${contactHeaders}<br>---${END_HEADERS}---</div>`;
        }
        const contacts = headers +
            `<div style="color: #AAAAAA;"><br>---${START_DETAILS}---<br>${contactDetails}<br>---${END_DETAILS}---</div>`;

        await setAsyncPromise((cb) => {
            Office.context.mailbox.item?.body.setAsync(currentData + contacts, {coercionType: Office.CoercionType.Html}, cb);
        }).catch(() => discard());
    }

    function discard() {
        document.getElementById("container")?.remove();

        const addContactButton = document.getElementById("addContact") as HTMLButtonElement;
        if (addContactButton) {
            addContactButton.classList.remove("ms-Button--disabled");
            addContactButton.disabled = false;
        } else {
            console.error("Something went very wrong with addContact button!");
        }

        const expContactButton = document.getElementById("exportContact") as HTMLButtonElement;
        if (expContactButton) {
            expContactButton.classList.remove("ms-Button--disabled");
            expContactButton.disabled = false;
        } else {
            console.error("Something went very wrong with expContact button!");
        }
    }

    async function exportContactData() {
        const eventTitle = await getAsyncPromise<string>((cb) =>
            Office.context.mailbox.item?.subject.getAsync(cb)
        ).catch(() => "");

        const bodyText = await getAsyncPromise<string>((cb) =>
            Office.context.mailbox.item?.body.getAsync(Office.CoercionType.Text, cb)
        ).catch(() => "");

        const pos = bodyText.indexOf(`---${START_AUTOMATED}---`);
        if (pos < 0) {
            console.error("No automated section found!");
            return;
        }
        const csv = bodyText
            .substring(pos, bodyText.length) // start after 1. disclaimer
            .replace(/(<br>)*---\[.*?]---(<br>)*/g, "") // remove disclaimer lines
            .replace(/\n+/g, "\n") // combine multiple consecutive linebreaks into one
            .replace(/^\s*$/gm, "");// remove empty lines
        const filename = (eventTitle || "event") + "_contacts";
        console.log(csv);
        downloadCSV(csv, filename);

    }

    function downloadCSV(csv: string, filename: string) {
        const blob = new Blob([csv], {type: 'text/csv;charset=utf-8;'});
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${filename}.csv`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    }

    function getAsyncPromise<T>(
        operation: (callback: (result: Office.AsyncResult<T>) => void) => void
    ): Promise<T> {
        return new Promise<T>((resolve, reject) => {
            operation((result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve(result.value);
                } else {
                    reject(new Error(result.error?.message ?? "Office get operation failed"));
                }
            });
        });
    }

    function setAsyncPromise(
        operation: (callback: (result: Office.AsyncResult<void>) => void) => void
    ): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            operation((result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve();
                } else {
                    reject(new Error(result.error?.message ?? "Office set operation failed"));
                }
            });
        });
    }

    function sanitizeInput(input: string): string {
        return input.trim().replace(/[<>\\;={}[\]$%"'`]/g, '');
    }
}
