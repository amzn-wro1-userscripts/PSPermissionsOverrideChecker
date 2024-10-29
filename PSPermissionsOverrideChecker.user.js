// ==UserScript==
// @name         PSPermissionsOverrideChecker
// @namespace    xdaaugus
// @version      1.0.0
// @description  Sprawdzanie uprawnień solverów do Grass, Denali, TCDA i innych magazynów w FC Research
// @author       Dawid Augustyn - xdaaugus
// @match        https://permissions.amazon.com/*
// @icon         https://cdn.pixabay.com/photo/2017/11/20/14/49/basket-2965747_1280.png
// @connect      amazon.com
// @connect      chime.aws
// @connect      quip-amazon.com
// @grant        GM.xmlHttpRequest
// @grant        unsafeWindow
// @grant        GM_addStyle
// @grant        window.onurlchange
// @grant        GM_registerMenuCommand
// @require      https://cdn.jsdelivr.net/npm/sweetalert2@11.7.3/dist/sweetalert2.all.min.js
// @require      https://unpkg.com/write-excel-file@1.x/bundle/write-excel-file.min.js
// ==/UserScript==

/*
	Coś nie działa? Chcesz dodać jakąś grupę uprawnień?
	Odezwij się do mnie na Chime - xdaaugus
	todo zmienic alerty na sweetalert2
*/

async function getPermissionsCSRFTokenFromHttpRequest() {
	return new Promise((resolve, reject) => {
		GM.xmlHttpRequest({
			method: 'GET',
			url: `https://permissions.amazon.com/a/`,
			onload: function(response){
				if (response.status !== 200) reject(response.status);
				const data = response.responseText;
				const dom = new DOMParser().parseFromString(data, 'text/html');
				const csrfToken = dom.querySelector('meta[name="csrf-token"]')?.content;
				if (!csrfToken) reject('No CSRF token found');
				resolve(csrfToken);
			}
		});
	});
}

let permissionsCache = {};
async function isEmployeeInPermissionsGroup(login, teamID, forceHttpRequest = false) { // bardziej technicznie to isEmployeeAddedAsOverrideInPermissionGroup ale nazwa za dluga
	if (!login || !teamID) return false;
	const storageID = `permissionsOverride_${teamID}`;
	if (forceHttpRequest) {
		const token = await getPermissionsCSRFTokenFromHttpRequest();
		return new Promise((resolve, reject) => {
			GM.xmlHttpRequest({
				url: 'https://permissions.amazon.com/a/team/index/search_team_overrides',
				method: 'POST',
				headers: {
					'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
					'X-CSRF-Token': token,
				},
				data: `team_id=${teamID}&token_input%5Bid%5D=${login}&token_input%5Bname%5D=&type=ADD&allow_action=false`,
				onload: async function(response){
					if (response.status !== 200) reject(response.status);
					let data = response.responseText;
				//	console.log('data', data);
					try {
						data = JSON.parse(data);
						if (!data || data.error) {
							if (data?.error_message == 'Incorrect paramaters') { // tak, literowka w ich kodzie
								await new Promise((resolve, reject) => setTimeout(() => resolve(), 1000));
								return isEmployeeInPermissionsGroup(login, teamID, true);
							}
							return reject(false);
						}						
					} catch (e) {
						console.log('PARSING ERROR ', e, data);
						reject(e);
						return false;
					}
					const parser = new DOMParser();
					const DOM = parser.parseFromString(data.html, 'text/html');
					const stringWithData = DOM.querySelectorAll('script')[1].textContent;
					const hasPermissions = stringWithData.indexOf('DT_RowClass') > -1;
					if (!hasPermissions) {
						if (!permissionsCache[teamID]) permissionsCache[teamID] = {};
						permissionsCache[teamID][login] = false;
						setTimeout(() => {
							delete permissionsCache[teamID][login];
						}, 1000 * 60 * 30); // 30 minut cache
						return reject('Brak uprawnień');
					} else {
						const regex = /data = (\[[^\]]*\])/g;
						const matches = [...stringWithData.matchAll(regex)];
					//	console.log('matches', matches);
						if (matches && matches.length) {
							const match = matches[1][1];
						//	console.log('match', match);
							const json = JSON.parse(match)[0];
						//	console.log('json', json);
							const overrideReasonElement = parser.parseFromString(json.override_reason, 'text/html');
							const overrideReason = overrideReasonElement.querySelector('textarea').innerText.trim();
						//	console.log('overrideReason', overrideReason);
							const overrideString = json.permissions;
							let overrideType = overrideString.indexOf('Permanent') > -1 ? 'Permanent' : 'Temporary';
						//	console.log('overrideType', overrideType);
							if (overrideType === 'Temporary') {
								const datesRegex = /(\d{2}\/\d{2}\/\d{4})/g;
								var [startDate, endDate] = overrideString.match(datesRegex);
							//	console.log('startDate', startDate, 'endDate', endDate);
							}
							let ret = {
								overrideReason: overrideReason,
								overrideType: overrideType,
							}
							if (overrideType === 'Temporary') {
								ret.overrideStartDate = startDate;
								ret.overrideEndDate = endDate;
							}
							if (!permissionsCache[teamID]) permissionsCache[teamID] = {};
							permissionsCache[teamID][login] = ret;
							setTimeout(() => {
								delete permissionsCache[teamID][login];
							}, 1000 * 60 * 30); // 30 minut cache
							return resolve(ret);
						} else {
							if (!permissionsCache[teamID]) permissionsCache[teamID] = {};
							permissionsCache[teamID][login] = false;
							setTimeout(() => {
								delete permissionsCache[teamID][login];
							}, 1000 * 60 * 30); // 30 minut cache
							return reject('Brak uprawnień');
						}
					}
				}

			});
		});
	} else {
		if (!permissionsCache[teamID]) permissionsCache[teamID] = {};
		if (permissionsCache[teamID][login] === false) return false;
		if (permissionsCache[teamID][login]) return permissionsCache[teamID][login];
		return isEmployeeInPermissionsGroup(login, teamID, true);
	}
}
unsafeWindow.isEmployeeInPermissionsGroup = isEmployeeInPermissionsGroup;

async function checkEmployeesPermissionsForGroups(employeesList, teamsList) {
	let obj = [];
	for (const i in employeesList) {
		const employee = employeesList[i]; // login
		let completedRequests = 0;
		setStatusText(`Sprawdzanie ${employee} <br>(${parseInt(i) + 1}/${employeesList.length})`);
		function updateTextTillEnd() {
			const interval = setInterval(() => {
				setStatusText(`Sprawdzanie ${employee} (${+completedRequests+1}/${teamsList.length})<br>(${parseInt(i) + 1}/${employeesList.length})`);
				if (completedRequests >= teamsList.length) clearInterval(interval);
			}, 100);
			function cancel() {
				clearInterval(interval);
			}
			return { cancel };
		}
		const updateText = updateTextTillEnd();
		const employeePromises = teamsList.map(async (teamName) => {
			const teamID = teams[teamName] ? teams[teamName][0] : false;
			let data = await isEmployeeInPermissionsGroup(employee, teamID).catch(e => false);
			completedRequests++;
			return { [teamName]: data };
		});
		let ret = { employee: employee, teams: {} };
		const employeeInPPR = await searchInPPR(employee).catch(e => false);
		if (employeeInPPR) {
			const employeeName = formatName(employeeInPPR.employeeName);
			const managerName = formatName(employeeInPPR.supervisorName);
			ret.employeeName = employeeName;
			ret.managerName = managerName;
		}
		let data = await Promise.all(employeePromises);
		updateText.cancel();
		data = data.reduce((acc, elem) => {
			return { ...acc, ...elem };
		}, {});
		console.log('data', data);
		ret.teams = data;
		obj.push(ret);
	}
//	console.log(obj);
	return obj;
}

const teams = {
	eufcresearch: ['amzn1.abacus.team.t6uxgwgxdp63ll3vvfga', 'EUFC-FCResearch', 'https://permissions.amazon.com/a/team/EUFC-FCResearch'],
	ac3: ['amzn1.abacus.team.63treezshzfu37k2gjsq', 'AC3', 'https://permissions.amazon.com/a/team/Raven_Extended_Research_Session2'],
    denali: ['amzn1.abacus.team.22ldgxhlyfd7crsj6htq', 'Denali', 'https://permissions.amazon.com/a/team/DenaliRetailManagerApprovedUsers'],
    grass: ['amzn1.abacus.team.iziua7ph77p3dca537xq', 'Grass', 'https://permissions.amazon.com/a/team/GreenerAccess_ldap_auto_migration_2024-03-18'],
    tcda: ['amzn1.abacus.team.pxm5q3iw6cbun6yn5xlq', 'TCDA', 'https://permissions.amazon.com/a/team/taxonomyaccesscontrolforgammaandprod_ldap_auto_migration_2023-07-07'], 
}
unsafeWindow.teams = teams;

////////////////////////  QUIP   ////////////////////////

function getAllowedCells(sheet) { // list of cells with data
	let cells = [];
	const range = (startColumn, startRow, endRow, description, place, extraIndexShiftData) => {
		for (let n = startRow; n <= endRow; n++) {
			let cellData = {
				position: `${startColumn}${n}`,
			};
			if (place) cellData.place = place;
			if (description) cellData.description = description;
			if (extraIndexShiftData) cellData.extraIndexShiftData = extraIndexShiftData;
			cells.push(cellData);
		}
	}
	if (sheet == 'QUIP Gold Urlopy') {
		range('A', 8, 150, 'Urlop');
	}
	return cells;
}


async function getStaticQuipData(appID, forceHttpRequest) {
	if (!appID) return false;
	const storageID = `quipStatic_${appID}`;
	if (forceHttpRequest) {
		return new Promise((resolve, reject) => {
			GM.xmlHttpRequest({
				method: 'GET',
				url: `https://quip-amazon.com/-/html/${appID}`,
				onload: function(response){
					if (response.status !== 200) reject(response.status);
					const data = response.responseText;
					localStorage[storageID] = JSON.stringify({ cacheTime: Date.now(), data: data });
					resolve(data);
				}
			});
		});
	} else {
		const storage = localStorage[storageID] ? JSON.parse(localStorage[storageID]) : false;
		if (storage) {
			if ( storage.cacheTime + 60000 < Date.now() ) // 60s cache
				return getStaticQuipData(appID, true); 
			return storage.data;
		} else return getStaticQuipData(appID, true);
	}
}

function numberToLetters(num) {
	num--;
	let letters = ''
	while (num >= 0) {
		letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[num % 26] + letters
		num = Math.floor(num / 26) - 1;
	}
	return letters
}

function qStatic_getAllEmployeesCells(documentName, table) {
	let ret = [];
	if (!documentName) return false;
	const allowedCellsWithFunctions = getAllowedCells(documentName);
	const rows = table.querySelectorAll('tbody tr');
	for (const rowIndex in rows) {
		const row = rows[rowIndex];
	//	console.log('rowIndex', rowIndex, 'row', row)
		if (!row.querySelector) continue;
		const rowCells = row.querySelectorAll('td');
		for (const cellIndex in rowCells) {
			const cell = rowCells[cellIndex];
		//	console.log('cellIndex', cellIndex, 'cell', cell)
			if (!cell.querySelector) continue;
			const cellName = `${numberToLetters(parseInt(cellIndex))}${parseInt(rowIndex) + 1}`;
		//	console.log(cellName)
			const allowedCellData = allowedCellsWithFunctions.find(cell => cell.position == cellName);
			if (!allowedCellData) continue;
			const login = cell.textContent.trim();
			if (!login || login == '' || login == ' ') continue;
			const extraIndexShiftData = allowedCellData.extraIndexShiftData;
			let extraData = {};
			if (extraIndexShiftData) {
				for (const extraDataName in extraIndexShiftData) {
					const indexShift = parseInt(extraIndexShiftData[extraDataName]);
					const extraCell = rowCells[+cellIndex + indexShift];
				//	console.log(`extraCell (${extraDataName} ${indexShift}: ind ${+cellIndex + indexShift})`, extraCell)
					const extraCellText = extraCell ? extraCell.textContent.trim() : false;
					if (!extraCellText || extraCellText == '' || extraCellText == ' ') continue;
					extraData[extraDataName] = extraCellText;
				}
			}
			if (allowedCellData.place) ret.push({ element: cell, text: login, position: cellName, description: allowedCellData.description, place: allowedCellData.place, ...extraData });
			else ret.push({ element: cell, text: login, position: cellName, description: allowedCellData.description, ...extraData });
		}
	}
	return ret;
}

const documentIDNames = {
	['9jzmAWR6FacG']: { name: 'QUIP Gold Urlopy', appID: 'Tdb9BAAem7L', tableIndex: 0 },
}

async function getAllEmployeesCells_QUIP(documentName = 'QUIP Gold Urlopy', forceHttpRequest = false) {
	const documentData = Object.values(documentIDNames).find(entry => entry.name == documentName);
	if (!documentData) return false;
	const appID = documentData.appID;
	const staticQuipData = await getStaticQuipData(appID, forceHttpRequest).catch(err => false);
	if (!staticQuipData) return false;
	const parser = new DOMParser();
	const htmlDoc = parser.parseFromString(staticQuipData, 'text/html');
	const tables = [...htmlDoc.querySelectorAll('table')];
	const table = tables[documentData.tableIndex || 0];
	return qStatic_getAllEmployeesCells(documentName, table);
}
unsafeWindow.getAllEmployeesCells_QUIP = getAllEmployeesCells_QUIP;

function formatName(name) {
	// Zielonka,Maciej Roch -> Maciej Zielonka
	const unformatName = name.split(',');
	return `${unformatName[1].split(' ')[0]} ${unformatName[0]}`;
}

async function searchInPPR(term, forceHttpRequest) {
	const isTermNumber = !isNaN(term);
	if (forceHttpRequest) {
		return new Promise((resolve, reject) => {
			GM.xmlHttpRequest({
				method: 'GET',
				url: `https://fclm-portal.amazon.com/ajax/partialEmployeeSearch?term=${term}`,
				onload: function(response){
					if (response.status !== 200) return reject(response.status);
					if (response.finalUrl.includes('midway-auth')) return reject('midway');
					let data = response.responseText;
					data = JSON.parse(data);
					if (data?.value?.length > 0) {
					//	data = data.value[0];
						data = data.value.find(employee => employee.employeeLogin == term || employee.badgeBarcodeId == term || employee.employeeId == term || isTermNumber && (parseInt(employee.badgeBarcodeId) == parseInt(term) || parseInt(employee.employeeId) == parseInt(term)));
					} else {
						resolve(false);
						return false;
					}
					if (!data) {
						resolve(false);
						return false;
					}
					let storage = sessionStorage.employeePPRSearchData ? JSON.parse(sessionStorage.employeePPRSearchData) : {};
					storage[term] = { cacheTime: Date.now(), data };
					sessionStorage.employeePPRSearchData = JSON.stringify(storage);
					resolve(data);
				}
			});
		});
	} else {
		const storage = sessionStorage.employeePPRSearchData ? JSON.parse(sessionStorage.employeePPRSearchData) : false;
		if (storage && storage[term]) {
			const data = storage[term];
			if ( data.cacheTime + 60000 * 60 * 24 * 7 < Date.now() ) {
				return searchInPPR(term, true); // 7 days cache
			}
			return data.data;
		} else return searchInPPR(term, true);
	}
}
unsafeWindow.searchInPPR = searchInPPR;

/////////////////////////////////////////////////////////

GM_registerMenuCommand('Check Problem Solvers permissions', async function(event) {
	showWindow();
});

function showWindow() {
	const body = document.querySelector('body');
	if (!body) return setTimeout(showWindow, 100);
	const checkboxesData = Object.entries(teams).map(entry => {
		return { id: entry[0], name: entry[1][1] };
	});
	const div = document.createElement('div');
	div.id = 'permissionsCheckerWindow';
	div.innerHTML = `
		<div id='dataEntry'>
			<h1>Problem Solve Permissions Override Checker</h1>
		<!--	<label for="destinationCode">Destination code</label>
			<input type="text" id="destinationCode" placeholder="Destination code" /> -->
			<div class='checkboxes'>
				${checkboxesData.map(checkbox => `
					<input type="checkbox" id="${checkbox.id}" name="${checkbox.name}" checked>
					<label for="${checkbox.id}">${checkbox.name}</label>
				`).join('')}
			</div>
			<p>Wklej loginy solverów do sprawdzenia, oddzielone enterem:<br>Lub naciśnij przycisk "Pobierz listę solverów z Quipa"</p>
			<textarea id='loginsInput' rows='15' spellcheck="false"></textarea>
			<div class='buttons'>
				<button id="fillProblemSolvers">Pobierz listę solverów z Quipa</button>
				<button id="checkPermissions">Sprawdź uprawnienia</button>
			</div>
		</div>
		<div id='progress'>
				<div class='title'>Status</div>
				<span id='currentstatus'>Nie uruchomiono</span>
			</div>
		</div>
		<div id='results' class='hidden'>
			<table id='resultsTable'>
				<thead>
					<tr>
						<th>Container</th>
						<th>Destination</th>
						<th>Data</th>
					</tr>
				</thead>
				<tbody></tbody>
			</table>
		</div>
	`;
	body.appendChild(div);
	body.style.overflow = 'hidden';

	document.getElementById('fillProblemSolvers').addEventListener('click', async () => {
		const logins = await getProblemSolversList();
		if (!logins) return;
		document.getElementById('loginsInput').value = logins.join('\n');
	});

	document.getElementById('checkPermissions').addEventListener('click', async () => {
		const logins = getLoginsFromInput();
		if (logins.length == 0) return alert('Wprowadź loginy solverów do sprawdzenia');
		let selectedCheckboxes = [];
		const checkboxes = document.querySelectorAll('.checkboxes input');
		for (const checkbox of checkboxes) {
			if (checkbox.checked) selectedCheckboxes.push(checkbox.id);
		}
		if (selectedCheckboxes.length == 0) return alert('Wybierz przynajmniej jedną grupę uprawnień do sprawdzenia');
		checkProblemSolversPermissions(logins, selectedCheckboxes);
	});
}

GM_addStyle(`
	#permissionsCheckerWindow {
		position: fixed;
		top: 0;
		left: 0;
		width: 100%;
		height: 100%;
		background: rgba(0, 0, 0, 0.9);
		backdrop-filter: blur(8px);
		color: white;
		display: flex;
		justify-content: center;
		align-items: center;
		z-index: 999;
	}

	#permissionsCheckerWindow h1 {
		color: crimson;
	}

	#permissionsCheckerWindow > #dataEntry {
		display: flex;
		flex-direction: column;
		align-items: center;
		gap: 0.2em;
	}
	
	#permissionsCheckerWindow > #dataEntry > p {
		margin: 0;
	}
	
	#permissionsCheckerWindow .checkboxes {
		display: flex;
		flex-direction: row;
		gap: 0.5em;
		align-items: last baseline;
	}

	#loginsInput {
		color: black;
		font-size: 105%;
		width: 50%;
		border-radius: 14px;
		margin-bottom: 0.5em;
	}

	#results.hidden {
		display: none;
	}
	
	#progress {
		position: absolute;
		top: 1em;
		right: 1em;
		background: cadetblue;
		padding: 1em;
		border-radius: 20px;
		min-width: 15em;
		text-align: center;
		transition: all 1s;
	}
	
	#progress .title {
		font-weight: bold;
		font-size: large;
	}
		
`);


function getLoginsFromInput() {
	const loginsInput = document.getElementById('loginsInput');
	const logins = loginsInput.value.split('\n').map(entry => entry.trim()).filter(entry => entry);
	// remove duplicates:
	return [...new Set(logins)];
}

function setStatusText(text) {
	document.getElementById('currentstatus').innerHTML = text;
}


/////////////////////////////////////////////////////////


async function getProblemSolversList() {
	const allEmployeesCells = await getAllEmployeesCells_QUIP('QUIP Gold Urlopy');
	if (!allEmployeesCells || allEmployeesCells.length == 0) {
		console.error('❌ Wystąpił błąd przy pobieraniu listy solverów z Quipa urlopowego');
		alert('❌ Wystąpił błąd przy pobieraniu listy solverów z Quipa urlopowego - upewnij się, że jesteś zalogowany w Quipie i masz dostęp do dokumentu "Urlopy Gold"');
		return false;
	}
	let logins = allEmployeesCells.map(employee => employee.text);
	logins = logins.sort((a, b) => a < b ? -1 : 1);
	return logins;
}
unsafeWindow.getProblemSolversList = getProblemSolversList;

function prepareEmployeesSummarySheetData(temporaryPermissionsData) {
	let headers = ['Login', 'Name', 'Manager'];
	let columnsSize = [];
	const firstEntry = temporaryPermissionsData[0];
	const checkedTeams = Object.keys(firstEntry.teams);
	for (const checkedTeam of checkedTeams) {
		const name = teams[checkedTeam][1];
		headers.push(name);
	}
	const HEADER_ROW = headers.map(header => {
		if (header == 'Login') columnsSize.push({ width: 15 });
		else if (header == 'Name' || header == 'Manager') columnsSize.push({ width: 20 });
		else columnsSize.push({ width: 12 });
		return { value: header, fontWeight: 'bold' };
	});
	const DATA_ROWS = temporaryPermissionsData.map(entry => {
		let row = [{ type: String, value: entry.employee }];
		if (entry.employeeName) {
			row.push({ type: String, value: entry.employeeName });
		} else {
			row.push({ type: String, value: 'N/A' });
		}
		if (entry.managerName) {
			row.push({ type: String, value: entry.managerName });
		} else {
			row.push({ type: String, value: 'N/A' });
		}
		for (const checkedTeam of checkedTeams) {
			const data = entry.teams[checkedTeam];
			if (data) row.push({ type: String, value: 'YES', backgroundColor: '#00ff00' });
			else row.push({ type: String, value: 'NO', backgroundColor: '#ff0000' });
		}
		return row;
	});
	const sheetData = [HEADER_ROW, ...DATA_ROWS];
	console.log('sheetData', sheetData);
	return { columns: columnsSize, sheetData };
}

function preparePermissionSheetData(temporaryPermissionsData, permissionName) {
	const headers = ['Login', 'Name', 'Manager', 'Override Type', 'Override Reason', 'Start Date', 'End Date'];
	let columnsSize = [];
	const HEADER_ROW = headers.map(header => {
		if (header == 'Login') columnsSize.push({ width: 15 });
		else if (header == 'Name' || header == 'Manager') columnsSize.push({ width: 20 });
		else columnsSize.push({ width: Math.ceil(+header.length*1.5) });
		return { value: header, fontWeight: 'bold' };
	});
	const DATA_ROWS = temporaryPermissionsData.map(entry => {
		let row = [{ type: String, value: entry.employee }];
		if (entry.employeeName) {
			row.push({ type: String, value: entry.employeeName });
		} else {
			row.push({ type: String, value: 'N/A' });
		}
		if (entry.managerName) {
			row.push({ type: String, value: entry.managerName });
		} else {
			row.push({ type: String, value: 'N/A' });
		}
		// teams
		const data = entry.teams[permissionName];
		if (data) {
			row.push({ type: String, value: data.overrideType });
			row.push({ type: String, value: data.overrideReason || 'N/A' });
			if (data.overrideType === 'Temporary') {
				row.push({ type: String, value: data.overrideStartDate || 'Error'});
				row.push({ type: String, value: data.overrideEndDate || 'Error'});
			} else {
				row.push({ type: String, value: 'N/A' });
				row.push({ type: String, value: 'N/A' });
			}
		} else {
			row.push({ type: String, value: 'None' });
			row.push({ type: String, value: 'N/A' });
			row.push({ type: String, value: 'N/A' });
			row.push({ type: String, value: 'N/A' });
		}
		return row;
	});
	const sheetData = [HEADER_ROW, ...DATA_ROWS];
	console.log('sheetData', sheetData);
	return { columns: columnsSize, sheetData };
}

async function checkProblemSolversPermissions(logins, permissions) {
	setStatusText('Rozpoczynanie...');
	const temporaryPermissions = await checkEmployeesPermissionsForGroups(logins, permissions);
	console.log('temporaryPermissions', temporaryPermissions);
	setStatusText('Generowanie raportu...');
	const summarySheetData = prepareEmployeesSummarySheetData(temporaryPermissions);
	const permissionSheetsData = permissions.map(permission => {
		return preparePermissionSheetData(temporaryPermissions, permission);
	});
	function getCurrentTime() {
		const date = new Date();
		const year = date.getFullYear();
		const month = ("0" + (date.getMonth() + 1)).slice(-2);
		const day = ("0" + date.getDate()).slice(-2);
		const hours = ("0" + date.getHours()).slice(-2);
		const minutes = ("0" + date.getMinutes()).slice(-2);
		const seconds = ("0" + date.getSeconds()).slice(-2);
		return `${year}-${month}-${day}_${hours}-${minutes}-${seconds}`;
	}
	const fileName = `${getCurrentTime()}_PS_temporary_permissions.xlsx`;
	console.log([summarySheetData.sheetData, ...permissionSheetsData.map(sheet => sheet.sheetData)]);
	console.log([summarySheetData, ...permissionSheetsData])
	await writeXlsxFile([summarySheetData.sheetData, ...permissionSheetsData.map(sheet => sheet.sheetData)], {
		fileName: fileName,
		columns: [summarySheetData.columns, ...permissionSheetsData.map(sheet => sheet.columns)],
		sheets: ['Summary', ...permissions.map(permission => teams[permission][1])]
	});

	setStatusText('Zakończono');
}
unsafeWindow.checkProblemSolversPermissions = checkProblemSolversPermissions;