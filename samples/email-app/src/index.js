function isLoggedIn() {
  const provider = mgt.Providers.globalProvider;
  return provider && provider.state === mgt.ProviderState.SignedIn;
}

function showStartPage() {
  if (isLoggedIn()) {
    const columns = getColumns();
    const main = document.getElementById('main-container');
    const newDiv = document.createElement('div');
    newDiv.setAttribute('class', 'container-fluid');
    newDiv.innerHTML = columns;
    main.replaceWith(newDiv);

    const accountNav = document.getElementById('account-nav');
    accountNav.innerHTML = '<mgt-login class="mgt-dark"></mgt-login>';

    const emailList = document.getElementById('email-list');
    emailList.innerHTML = getEmails();

    const rightCol = document.getElementById('right-col');
    rightCol.innerHTML = getTabs();

    const navList = document.getElementById('nav-list');
    const listItem = document.createElement('li');
    listItem.setAttribute('class', 'nav-item');
    listItem.innerHTML = `
      <a class="nav-link" href="#" onclick="showEmailComposer()">
        <i class="fas fa-pen-alt mr-1"></i> Compose New Email</a>`;
    navList.appendChild(listItem);
  } else {
    const main = document.getElementById('main-container');
      console.log(main);
    main.innerHTML = `
    <div class="jumbotron text-center">
      <h1>Contoso Mail</h1>
      <p class="lead">Send awesome spam emails to your friends</p>
      <mgt-login></mgt-login>
    </div>`;
  }
}

function getColumns() {
  const container = `<div class="row justify-content-center"><!-- Dashboard --></div>`;

  const dashboard = `
    <div class="col-3 bg-light main-sidebar main-left-sidebar">
      <div class="sidebar" id="email-list"></div>
    </div>
    <main class="col-6">
      <div class="sidebar" id="email-body">${getEmailForm()}</div>
    </main>
    <div class="col-3 bg-light main-sidebar main-right-sidebar">
      <div class="sidebar" id="right-col"></div>
    </div>`;
  return container.replace('<!-- Dashboard -->', dashboard);
}

function getEmails() {
  return `
    <mgt-get resource="/me/messages" version="beta" scopes="mail.read" max-pages="2">
        <template>
            <ul class="nav flex-column sidebar-nav" data-for="email in value">
                <li class="nav-item" id="{{email.id}}" onclick=showEmail({{email}})>
                  <a class="nav-link" href="#">{{ email.subject }}</a>
                </li>
            </ul>
        </template>
        <template data-type="loading">loading emails</template>
        <template data-type="error">{{ this }}</template>
    </mgt-get>`;
}

function showEmail(email) {
  const emailBody = email.body.content;
  const emailBodyDiv = document.getElementById('email-body');
  emailBodyDiv.innerHTML = emailBody;
}

function getTabs() {
  return `
  <nav>
    <div class="nav nav-tabs" id="nav-tab" role="tablist">
      <button class="nav-link active" id="nav-agenda-tab" data-bs-toggle="tab" data-bs-target="#nav-agenda" type="button" role="tab" aria-controls="nav-agenda" aria-selected="true">My agendas</button>
      <button class="nav-link" id="nav-tasks-tab" data-bs-toggle="tab" data-bs-target="#nav-tasks" type="button" role="tab" aria-controls="nav-tasks" aria-selected="false">My tasks</button>
      <button class="nav-link" id="nav-files-tab" data-bs-toggle="tab" data-bs-target="#nav-files" type="button" role="tab" aria-controls="nav-files" aria-selected="false">My files</button>
    </div>
  </nav>
  <div class="tab-content" id="nav-tabContent">
    <div class="tab-pane fade show active" id="nav-agenda" role="tabpanel" aria-labelledby="nav-agenda-tab">
      <br>
      <mgt-agenda></mgt-agenda>
    </div>
    <div class="tab-pane fade" id="nav-tasks" role="tabpanel" aria-labelledby="nav-tasks-tab">
      <mgt-tasks read-only>
        <template data-type="task">
          <div class="card">
            <div class="card-body">
              <h5 class="card-title">{{task.title}}</h5>
              <h6 class="card-subtitle mb-2 text-muted">{{task.groupTitle}}: <small>{{task.folderTitle}}</small></h6>
              <div data-if="task.hasDescription">
                <mgt-get resource="/planner/tasks/{{task.id}}/details" scopes="group.read.all">
                  <template>
                    <p class="card-text">{{description}}</p>
                  </template>
                </mgt-get>
              </div>
              <div data-else>
                <p class="card-text text-muted"><i>No description</i></p>
              </div>
              <div>
                <p class="card-text small">Due: {{readableDate(task.dueDateTime)}}</p>
                <div class="progress">
                  <div
                    class="progress-bar"
                    role="progressbar"
                    style="width: {{formatNumber(task.percentComplete)}}%;"
                    aria-valuenow="{{formatNumber(task.percentComplete)}}"
                    aria-valuemin="0" aria-valuemax="100">
                      {{formatNumber(task.percentComplete)}}%
                  </div>
                </div>
              </div>
              <br>
              <mgt-person person-query="{{task.createdBy.user.id}}" person-card="hover"></mgt-person>
            </div>
          </div>
          <br>
        </template>
      </mgt-tasks>
    </div>
    <div class="tab-pane fade" id="nav-files" role="tabpanel" aria-labelledby="nav-files-tab">
      ${getOnedriveFiles()}
    </div>
  </div>`;
}

function getEmailForm() {
  return `
    <form>
      <div class="mb-3">
        <label for="emailSubject" class="form-label">Email subject</label>
        <input type="text" class="form-control" id="emailSubject" placeholder="The subject is..." required>
      </div>
      <mgt-people-picker required></mgt-people-picker>
      <div class="mb-3">
        <label for="emailBody" class="form-label">Email body</label>
        <textarea class="form-control" id="emailBody" rows="3" required></textarea>
      </div>
      <button type="button" class="btn btn-primary" onclick="sendEmail()">Send</button>
    </form>`;
}

function showEmailComposer() {
  const emailComposer = getEmailForm();
  const emailBodyNode = document.getElementById('email-body');
  emailBodyNode.innerHTML = emailComposer;
}

async function sendEmail() {
  let selectedPeople = document.querySelector('mgt-people-picker').selectedPeople;
  const emailSubject = document.getElementById('emailSubject').value;
  const emailBody = document.getElementById('emailBody').value;
  const form = document.querySelector('form');
  if (selectedPeople.length) {
    const recipients = [];
    for (let i = 0; i < selectedPeople.length; i++) {
      const person = selectedPeople[i];
      const emailObj = {
        emailAddress: {
          address: person.userPrincipalName
        }
      };
      recipients.push(emailObj);
    }
    const emailObject = {
      subject: emailSubject,
      body: {
        contentType: 'Text', // NOTE: Currently sending text alone.
        content: emailBody
      },
      toRecipients: recipients
    };
    const provider = mgt.Providers.globalProvider;
    if (provider) {
      const graphClient = provider.graph.client;
      try {
        const sentEmailResponse = await graphClient.api('me/sendMail').post({ message: emailObject });
        console.log(sentEmailResponse);
      } catch (error) {
        console.error(error);
      }
      form.reset();
      document.querySelector('mgt-people-picker').selectedPeople = [];
    }
  } else {
    alert('Add recipients!');
    form.reset();
    document.querySelector('mgt-people-picker').selectedPeople = [];
    return;
  }
}

function getOnedriveFiles() {
  return `
    <mgt-get resource="/me/drive/root/children" scopes="files.read.all">
      <template>
      <br>
        <ul class="list-group flex-column sidebar-nav" data-for="file in value">
          <li class="list-group-item d-flex justify-content-between align-items-start">
            <div class="ms-2 me-auto">
              <div class="fw-bold"><a href="{{file.webUrl}}">{{file.name}}</a></div>
              <span><small> Created by: </small><mgt-person
                person-query="{{file.createdBy.user.email}}"
                view="oneline"
                person-card="hover"></mgt-person></span>
            </div>
            <span class="badge bg-primary rounded-pill">{{formatFileSize(file.size)}}</span>
          </li>
          <br>
        </ul>
      </template>
      <template data-type="loading">
        loading files
      </template>
      <template data-type="error">
        {{ this }}
      </template>
    </mgt-get>`;
}

function readableDate(date) {
  return new Date(date).toDateString();
}

function formatFileSize(bytes, decimalPoint) {
  if (bytes == 0) return '0 Bytes';
  var k = 1000,
    dm = decimalPoint || 2,
    sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'],
    i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

function formatNumber(number) {
  if (number === 0) {
    return 1;
  }
  return number;
}
showStartPage();
