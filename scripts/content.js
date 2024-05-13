const table = document.getElementById('rubric-criteria')
const tbody = table.lastElementChild
const addCriterionButton = document.getElementById('rubric-criteria-addcriterion')
const rubric = document.getElementById('fitem_id_rubric')

function selectFile(contentType, multiple = false) {
    return new Promise(resolve => {
        const input = document.createElement('input')
        input.type = 'file'
        input.multiple = multiple
        input.accept = contentType

        input.addEventListener('change', () => {
            const files = Array.from(input.files)
            if (multiple) {
                resolve(files)
            } else {
                resolve(files[0])
            }
        })

        input.click()
    })
}

function getNextKey(key) {
    if (key === 'Z' || key === 'z') {
        return String.fromCharCode(key.charCodeAt(0) - 25) + String.fromCharCode(key.charCodeAt(0) - 25) // AA or aa
    }

    const lastChar = key.slice(-1)
    const sub = key.slice(0, -1)
    if (lastChar === 'Z' || lastChar === 'z') {
        // If a string of length > 1 ends in Z/z,
        // increment the string (excluding the last Z/z) recursively,
        // and append A/a (depending on casing) to it
        return getNextKey(sub) + String.fromCharCode(lastChar.charCodeAt(0) - 25)
    }
    // (take till last char) append with (increment last char)

    return sub + String.fromCharCode(lastChar.charCodeAt(0) + 1)


}

function clearForm() {
    let count = tbody.children.length
    for (let i = 0; i < count; i++) {
        const tr = tbody.firstElementChild
        const controls = tr.firstElementChild
        const del = controls.children[1].firstElementChild
        del.click()
    }
}

function newCriterion() {
    addCriterionButton.click()
}

function modifyCriterion(criterion, name) {
    // Sets criterion name
    const tr = tbody.children[criterion]
    const textarea = tr.getElementsByClassName('description')[0].firstElementChild
    textarea.parentElement.click()
    textarea.value = name
    textarea.blur()
}

function newLevelInCriterion(criterion) {
    const tr = tbody.children[criterion]
    const addLevelButton = tr.getElementsByClassName('addlevel')[0].firstElementChild
    addLevelButton.click()
}

function modifyLevelInCriterion(criterion, level, description, grade) {
    const tr = tbody.children[criterion]
    const levels = tr.getElementsByTagName('table')[0]
    const tbodyLevels = levels.firstElementChild
    const levelChild = tbodyLevels.firstElementChild.children[level]

    levelChild.click()

    const levelTextarea = levelChild.getElementsByClassName('definition')[0].firstElementChild
    levelTextarea.value = description

    const gradeInput = levelChild.getElementsByClassName('score')[0].firstElementChild.firstElementChild
    gradeInput.value = grade
    gradeInput.blur()
}

function processExcel(data, offset) {
    const opts = {type: 'binary'}
    const workbook = XLSX.read(data, opts)
    const sheet = workbook.Sheets[workbook.SheetNames[0]]

    const defaultCriteria = 1
    let defaultLevels = 3
    // Excel indexes
    let row = 1 + offset
    let criterionCell = 'A' + row
    // Moodle indexes
    let criterion = 0

    do {
        // Excel indexes
        let levelColumn = 'B'
        let levelCell = levelColumn + row
        let gradeCell = levelColumn + (row + 1)
        // Moodle indexes
        let level = 0

        if (criterion >= defaultCriteria) {
            newCriterion()
        }
        modifyCriterion(criterion, sheet[criterionCell].v)

        do {
            if (level >= defaultLevels) {
                newLevelInCriterion(criterion)
            }
            modifyLevelInCriterion(criterion, level, sheet[levelCell].v, sheet[gradeCell].v)

            levelColumn = getNextKey(levelColumn)

            levelCell = levelColumn + row
            gradeCell = levelColumn + (row + 1)
            level++
        } while (sheet[levelCell])

        row += 2
        criterionCell = 'A' + row
        criterion++
        // Moodle by default will create new criteria with as many levels as the previous maximum
        defaultLevels = Math.max(defaultLevels, level)
    } while (sheet[criterionCell])
}

async function onImportClick(e) {
    e.preventDefault()

    const input = prompt("Row offset", "1")
    if (input == null) {
        return
    }

    const offset = parseInt(input, 10)

    if (isNaN(offset) || offset < 0) {
        alert("Error: The row offset must be a positive integer")
        return
    }

    const file = await selectFile('.xls,.xlsx')

    const reader = new FileReader()
    reader.addEventListener('load', (event) => {
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
