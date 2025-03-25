let apiKey;

Office.onReady(() => {
    console.log('Office.js is ready');
    initializeAddIn();
}).catch(error => {
    console.error('Error initializing Office:', error);
});

function initializeAddIn() {
  console.log('Office is ready');
  retrieveApiKey();
  document.getElementById('saveApiKey').addEventListener('click', saveApiKey);
  document.getElementById('workspaceSelect').addEventListener('change', onWorkspaceSelect);
  document.getElementById('projectSelect').addEventListener('change', onProjectSelect);
  document.getElementById('addNewTaskButton').addEventListener('click', () => {
        document.getElementById('newTaskForm').style.display = 'block';
    });
  document.getElementById('saveNewTask').addEventListener('click', handleSaveNewTask);
  document.getElementById('createTimeEntry').addEventListener('click', createTimeEntry);
  console.log(document.getElementById('toggleTheme'));
}

function showApiKeySection() {
    document.getElementById('settings').style.display = 'block';
}

// Save the API key
function saveApiKey() {
    apiKey = document.getElementById('apiKey').value; // Adjust based on your input
    localStorage.setItem('apiKey', apiKey);
    localStorage.removeItem('selectedWorkspaceId');
    console.log('API key saved to localStorage');
    showWorkspaceSelection();
    fetchWorkspaces(apiKey);
    setDefaultDescription();
}

// Retrieve the API key
function retrieveApiKey() {
    const savedKey = localStorage.getItem('apiKey');
    if (savedKey) {
        console.log('API key retrieved from localStorage:', savedKey);
        apiKey = savedKey; // Assign to global apiKey
        document.getElementById('apiKey').value = savedKey;
        showWorkspaceSelection();
        fetchWorkspaces(apiKey);
        setDefaultDescription();
        // Use the key in your add-in
    } else {
        console.log('No API key in localStorage');
        showApiKeySection();
    }

}

function showWorkspaceSelection() {
    document.getElementById('workspaceSelection').style.display = 'block';
}

function fetchWorkspaces(apiKey) {
    if (!apiKey) {
        showMessage('API key not set. Please enter and save your API key.');
        return;
    }
    fetch('https://api.clockify.me/api/v1/workspaces', {
        headers: { 'X-Api-Key': apiKey }
    })
    .then(response => {
        if (!response.ok) throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        return response.json();
    })
    .then(data => {
        const workspaceSelect = document.getElementById('workspaceSelect');
        workspaceSelect.innerHTML = '<option value="">Select Workspace</option>';
        data.forEach(workspace => {
            const option = document.createElement('option');
            option.value = workspace.id;
            option.text = workspace.name;
            workspaceSelect.add(option);
        });

        const savedWorkspaceId = localStorage.getItem('selectedWorkspaceId');
        if (savedWorkspaceId && workspaceSelect.querySelector(`option[value="${savedWorkspaceId}"]`)) {
            workspaceSelect.value = savedWorkspaceId;
            onWorkspaceSelect({ target: workspaceSelect }); // Simulate the selection event
        }
    })
    .catch(error => {
        console.error('Error fetching workspaces:', error);
        showMessage(`Failed to fetch workspaces: ${error.message}`);
    });
}

function onWorkspaceSelect(event) {
    const workspaceId = event.target.value;
    if (workspaceId) {
        localStorage.setItem('selectedWorkspaceId', workspaceId);
        $('#projectTaskCollapse').collapse('show');
        $('#configCollapse').collapse('hide'); // Collapse configuration
        fetchProjects(workspaceId, apiKey);
        fetchTags(workspaceId, apiKey);
    }
}

function fetchProjects(workspaceId, apiKey, page = 1, allProjects = []) {
    fetch(`https://api.clockify.me/api/v1/workspaces/${workspaceId}/projects?page=${page}&page-size=50&archived=false`, {
        headers: { 'X-Api-Key': apiKey }
    })
    .then(response => response.json())
    .then(data => {
        const projects = allProjects.concat(data);
        if (data.length === 50) {
            fetchProjects(workspaceId, apiKey, page + 1, projects);
        } else {
            const projectSelect = document.getElementById('projectSelect');
            projectSelect.innerHTML = '<option value="">Select Project</option>';
            projects.forEach(project => {
                const option = document.createElement('option');
                option.value = project.id;
                option.text = project.name;
                projectSelect.add(option);
            });

						// Initialize Select2 for filterable dropdown
            $('#projectSelect').select2({
                placeholder: "Select or type to filter projects",
                allowClear: true // Adds an 'x' to clear the selection
            });

            // Attach event listener for project selection
            $('#projectSelect').on('select2:select', function (e) {
                onProjectSelect(e);
            });

        }
    })
    .catch(error => {
        console.error('Error fetching projects:', error);
        showMessage('Error fetching projects.');
    });
}

function onProjectSelect(event) {
    const projectId = event.params.data.id;
    const workspaceId = document.getElementById('workspaceSelect').value;
    if (projectId && workspaceId) {
        fetchTasks(workspaceId, projectId, apiKey);
    }
}

function fetchTasks(workspaceId, projectId, apiKey) {
    fetch(`https://api.clockify.me/api/v1/workspaces/${workspaceId}/projects/${projectId}/tasks`, {
        headers: { 'X-Api-Key': apiKey }
    })
    .then(response => response.json())
    .then(data => {
        const taskSelect = document.getElementById('taskSelect');
        taskSelect.innerHTML = '<option value="">Select Task</option>';
        data.forEach(task => {
            const option = document.createElement('option');
            option.value = task.id;
            option.text = task.name;
            taskSelect.add(option);
        });
    })
    .catch(error => {
        console.error('Error fetching tasks:', error);
        showMessage('Error fetching tasks.');
    });
}

function handleSaveNewTask() {
    const newTaskName = document.getElementById('newTaskName').value;
    const workspaceId = document.getElementById('workspaceSelect').value;
    const projectId = document.getElementById('projectSelect').value;

    // Basic validation
    if (!newTaskName || !projectId) {
        showMessage('Please enter a task name and select a project.');
        return;
    }

    // Use the global apiKey variable instead of hardcoding
    fetch(`https://api.clockify.me/api/v1/workspaces/${workspaceId}/projects/${projectId}/tasks`, {
        method: 'POST',
        headers: {
            'X-Api-Key': apiKey, // Assumes apiKey is a global variable set elsewhere
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ name: newTaskName })
    })
    .then(response => {
        if (!response.ok) throw new Error('Failed to create task');
        return response.json();
    })
    .then(newTask => {
        document.getElementById('newTaskForm').style.display = 'none';
        document.getElementById('newTaskName').value = '';
        fetchTasks(workspaceId, projectId, apiKey); // Refresh tasks with apiKey
        showNewTaskMessage('Task created successfully');
    })
    .catch(error => {
        console.error('Error:', error);
        showNewTaskMessage('Failed to create task. Please try again.');
    });
}

function fetchTags(workspaceId, apiKey) {
    fetch(`https://api.clockify.me/api/v1/workspaces/${workspaceId}/tags`, {
        headers: { 'X-Api-Key': apiKey }
    })
    .then(response => response.json())
    .then(data => {
        const tagSelect = document.getElementById('tagSelect');
        tagSelect.innerHTML = '';
        data.forEach(tag => {
            const option = document.createElement('option');
            option.value = tag.id;
            option.text = tag.name;
            tagSelect.add(option);
        });
    })
    .catch(error => {
        console.error('Error fetching tags:', error);
        showMessage('Error fetching tags.');
    });
}

function formatLocalDateTime(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0'); // Months are 0-based
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return `${year}-${month}-${day}T${hours}:${minutes}`;
}

function setDefaultDescription() {
  const appointment = Office.context.mailbox.item;
  if (appointment) {
    if (appointment.itemType === Office.MailboxEnums.ItemType.Appointment) {
      if (typeof appointment.start === 'object' && appointment.start.getAsync) {
        // Compose mode: Use asynchronous methods
        appointment.start.getAsync((startResult) => {
          if (startResult.status === Office.AsyncResultStatus.Succeeded) {
            const startDate = startResult.value;
            document.getElementById('startTime').value = formatLocalDateTime(startDate);
            // Set description
            appointment.subject.getAsync((subjectResult) => {
              if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
                document.getElementById('description').value = subjectResult.value || '';
              }
            });
          } else {
            console.error('Error getting appointment start time:', startResult.error);
            showMessage('Error retrieving appointment start time.');
          }
        });
        appointment.end.getAsync((endResult) => {
          if (endResult.status === Office.AsyncResultStatus.Succeeded) {
            const endDate = endResult.value;
            document.getElementById('endTime').value = formatLocalDateTime(endDate);
          } else {
            console.error('Error getting appointment end time:', endResult.error);
            showMessage('Error retrieving appointment end time.');
          }
        });
      } else {
        // Read mode: Use start and end directly
        const startDate = appointment.start;
        const endDate = appointment.end;
        document.getElementById('startTime').value = formatLocalDateTime(startDate);
        document.getElementById('endTime').value = formatLocalDateTime(endDate);
        document.getElementById('description').value = appointment.subject || '';
      }
    } else {
      showMessage('This add-in is only supported for appointments.');
    }
  } else {
    showMessage('No appointment context available. Please open this add-in from an appointment.');
  }
}

function createTimeEntry() {
  const workspaceId = document.getElementById('workspaceSelect').value;
  const projectId = document.getElementById('projectSelect').value;
  const taskId = document.getElementById('taskSelect').value;
  const description = document.getElementById('description').value;
  const tagIds = Array.from(document.getElementById('tagSelect').selectedOptions).map(option => option.value);
  const startTime = document.getElementById('startTime').value;
  const endTime = document.getElementById('endTime').value;

  // Validate required fields
  if (!workspaceId || !projectId || !description || !startTime || !endTime) {
    showMessage('Please select all required fields and enter a description, start time, and end time.');
    return;
  }

  const startDate = new Date(startTime);
  const endDate = new Date(endTime);

  if (isNaN(startDate) || isNaN(endDate)) {
    showMessage('Invalid start or end time.');
    return;
  }

    if (startDate >= endDate) {
        showMessage('End time must be after start time.');
        return;
    }

  const timeEntry = {
    description: description,
    end: endDate.toISOString(),
    projectId: projectId,
    start: startDate.toISOString(),
    tagIds: tagIds,
    taskId: taskId,
    billable: false,
    type: "REGULAR"
  };

  if (taskId && taskId.trim() !== '') {
    timeEntry.taskId = taskId;
  }

  fetch(`https://api.clockify.me/api/v1/workspaces/${workspaceId}/time-entries`, {
    method: 'POST',
    headers: {
      'X-Api-Key': apiKey,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(timeEntry)
  })
  .then(response => {
    if (response.ok) {
			$('#projectTaskCollapse').collapse('hide');
      showMessage('Time entry created successfully');
    } else {
      response.text().then(text => showMessage(`Failed to create time entry: ${text}`));
    }
  })
  .catch(error => showMessage('Error creating time entry: ' + error.message));
}

function showNewTaskMessage(message) {
    const newTaskMessageArea = document.getElementById('newTaskMessageArea');
    newTaskMessageArea.innerHTML = message;
    newTaskMessageArea.style.display = 'block';
}

function showMessage(message) {
    const messageArea = document.getElementById('messageArea');
    messageArea.innerHTML = message;
    messageArea.style.display = 'block';
}