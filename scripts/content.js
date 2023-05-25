const table = document.getElementById('rubric-criteria')
const addCriterionButton = document.getElementById('rubric-criteria-addcriterion')
const rubric = document.getElementById('fitem_id_rubric')

function selectFile(contentType, multiple=false) {
    return new Promise(resolve => {
        const input = document.createElement('input')
        input.type = 'file'
        input.multiple = multiple
        input.accept = contentType

        input.addEventListener('change', () => {
            const files = Array.from(input.files)
            if (multiple) {
                resolve(files)
            }
            else{
                resolve(files[0])
            }
        })

        input.click()
    })
}

function getNextKey(key) {
  if (key === 'Z' || key === 'z') {
    return String.fromCharCode(key.charCodeAt() - 25) + String.fromCharCode(key.charCodeAt() - 25) // AA or aa
  } else {
    const lastChar = key.slice(-1)
    const sub = key.slice(0, -1)
    if (lastChar === 'Z' || lastChar === 'z') {
      // If a string of length > 1 ends in Z/z,
      // increment the string (excluding the last Z/z) recursively,
      // and append A/a (depending on casing) to it
      return getNextKey(sub) + String.fromCharCode(lastChar.charCodeAt() - 25)
    } else {
      // (take till last char) append with (increment last char)
      return sub + String.fromCharCode(lastChar.charCodeAt() + 1)
    }
  }
  return key
}

function delay(timeout) {
	return new Promise(resolve => window.setTimeout(() => {
		resolve()
	}, timeout))
}

function confirm() {
	const confirm = document.getElementsByClassName('moodle-dialogue-confirm')[0]
	const yes = confirm.getElementsByTagName('input')[0]
	yes.click()
}

function clearForm() {
	const tbody = table.firstElementChild


	let count = tbody.children.length
	for (let i = 0; i < count; i++) {
		const tr = tbody.firstElementChild
		const controls = tr.firstElementChild
		const del = controls.children[1].firstElementChild
		del.click()

		confirm()
	}
}

function newCriterion(criterion) {
	addCriterionButton.click()

	const tr = table.firstElementChild.lastElementChild

	// Resets levels
	const levels = tr.getElementsByTagName('table')[0]
	const tbody = levels.firstElementChild

	const count = tbody.firstElementChild.children.length
	for (let i = 0; i < count; i++) {
		const level = tbody.firstElementChild.lastElementChild
		const del = level.firstElementChild.lastElementChild.firstElementChild
		del.click()
		confirm()
	}

	// Sets criterion name
	const textarea = tr.getElementsByClassName('description')[0].firstElementChild
	textarea.parentElement.click()
	textarea.value = criterion
	textarea.blur()
}

function addLevelToLastCriterion(description, grade) {
	const tr = table.firstElementChild.lastElementChild
	const addLevelButton = tr.getElementsByClassName('addlevel')[0].firstElementChild
	addLevelButton.click()

	const levels = tr.getElementsByTagName('table')[0]
	const tbody = levels.firstElementChild
	const level = tbody.firstElementChild.lastElementChild
	
	const levelTextarea = level.getElementsByClassName('definition')[0].firstElementChild
	levelTextarea.value = description
	
	const gradeInput = level.getElementsByClassName('score')[0].firstElementChild.firstElementChild
	gradeInput.value = grade
	gradeInput.blur()
}

function processExcel(data, offset) {
	const opts = { type: 'binary' }
	const workbook = XLSX.read(data, opts)
	const sheet = workbook.Sheets[workbook.SheetNames[0]]

	let row = 1 + offset
	let criterionCell = 'A' + row

	do {
		let levelColumn = 'B'
		let levelCell = levelColumn + row
		let gradeCell = levelColumn + (row + 1)

		newCriterion(sheet[criterionCell].v)

		do {
			addLevelToLastCriterion(sheet[levelCell].v, sheet[gradeCell].v)

			levelColumn = getNextKey(levelColumn)

			levelCell = levelColumn + row
			gradeCell = levelColumn + (row + 1)
		} while (sheet[levelCell])

		row += 2
		criterionCell = 'A' + row
	} while (sheet[criterionCell])
}

async function onImportClick(e) {
	e.preventDefault()

	const offset = parseInt(prompt("Row offset", "1"), 10)

	if (offset == NaN || offset < 0) {
		alert("Error: The row offset must be a positive integer")
		return
	}

	const file = await selectFile('.xls,.xlsx')

	const reader = new FileReader()
	reader.addEventListener('load', (event) => {
		clearForm()
		processExcel(event.target.result, offset)
	})
	reader.readAsBinaryString(file)
}

function createButton() {
	const wrapper = document.createElement('div')
	wrapper.classList.add('form-group')
	wrapper.classList.add('d-flex')
	wrapper.classList.add('flex-column')
	wrapper.classList.add('align-items-end')

	const button = document.createElement('button')
	button.classList.add('btn-plugin')
	button.classList.add('btn-plugin-excel')
	button.innerHTML = 'Import Excel'

	button.addEventListener('click', onImportClick)

	wrapper.append(button)

	return wrapper
}


rubric.prepend(createButton())
