function insertText() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rowI = 1;
  var allProjects = APIRequest('projects').projects;
  var totalRatingAvg = getTotalRatingAvg();
  var quantityRatings5 = getQuantityRatings(5);
  var quantityRatings4 = getQuantityRatings(4);
  var quantityRatings3 = getQuantityRatings(3);
  var quantityRatings2 = getQuantityRatings(2);

  OPTIONS.contractNumbers.forEach(function(contract, indexContract) {
    var regExpProject = new RegExp('^' + contract + '-[0-9]+-SUP');
    var projects = allProjects.filter(function(project) {
      return regExpProject.test(project.name);
    });
    var hoursSLA = parseFloat(APIRequestById('projects', projects[0].parent.id).project.custom_fields.find(function(i) {return i.id === 41}).value);
    hoursSLA = hoursSLA ? hoursSLA : 0;

    var textMail = 'Добрый день, коллеги!\n';
    textMail += 'Выражаем благодарность за длительное сотрудничество наших компаний. ';
    textMail += 'Мы очень рады нашему сотрудничеству и крайне признательны вам за лояльность и уважительное отношение к нашей работе. ';
    textMail += 'Во вложении прикреплён отчёт за ';
    textMail += stringMonth(OPTIONS.startDate) + ' ' + OPTIONS.startDate.getFullYear() + ' г.';
    textMail += ' по всем оказанным нами работам для Компании ';
    textMail += OPTIONS.companyName[indexContract];
    textMail += '. Средняя оценка ваших коллег нашей работы составляет ' + totalRatingAvg;
    textMail += ', что соответствует качественному параметру "' + stringRating(totalRatingAvg) + '". ';
    textMail += 'В этом месяце ваши коллеги поставили следующее количество оценок:\n';
    textMail += '5 - ' + quantityRatings5 + '\n';
    textMail += '4 - ' + quantityRatings4 + '\n';
    textMail += '3 - ' + quantityRatings3 + '\n';
    textMail += '2 - ' + quantityRatings2 + '\n\n';
    textMail += 'Задачи, оцененные на 2 и 3 не включаются в отчет, мы не берем за них деньги и не вычитаем включенные в договор часы. ';
    textMail += 'Они уже рассмотрены в качестве претензий и обрабатываются отделом контроля качества. ';
    textMail += 'По ним мы проводим работу и предпринимаем соответсвующие меры по улучшению качества нашей работы.\n\n';
    textMail += 'В Ваше комплексное обслуживание (______ руб/мес) номинально входит ' + hoursSLA + ' часов поддержки. ';
    textMail += 'Это количество часов складывается из расчёта сниженной стоимости поддержки для заказчика и минимальной оплаты за работу исполнителя (нашего сотрудника). ';
    textMail += 'Стоимость дальнейшей поддержки высчитывается по тарифу _____ руб/ч. В этом месяце для вас работали следующие инженеры нашей компании:\n';

    getWorkUsers(projects).forEach(function(user) {
      if (user.timeSpend > 0)
        textMail += user.name  + ' - ' + user.timeSpend + ' ч. Средняя оценка ' + user.ratingAvg + '\n';
    });

    var timeSpendByProjects = getTimeSpendByProjects(projects);
    textMail += '\nСогласно отчёту за предыдущий месяц, на работы было потрачено ' + timeSpendByProjects + '. ';
    textMail += 'Следовательно, максимально допустимое количество часов было превышено на ' + (timeSpendByProjects - hoursSLA);
    textMail += '. К сожалению, ввиду большого объема выполняемых работ, включать эти часы в счёт ежемесячного комплексного обслуживания. ';
    textMail += 'В связи с этим, выставляем счёт на сумму _______ рублей.\n\n';
    textMail += 'Также считаем нужным в дальнейшем прописать этот момент в дополнительном соглашении к договору (во вложении). ';
    textMail += 'Ранее же эта информация была доступна в личном кабинете заказчика, который указан в договоре.\n\n';
    textMail += 'В случае каких-либо вопросов или несогласия, можно обсудить эту тему с нашим коммерческим директором Алиной Муракаевой (тел. ...), ';
    textMail += 'либо с генеральным директором Дмитрием Сабуровым (тел. ...).';

    sheet.getRange(rowI, 1).setValue(textMail);

    rowI++;
  });
}

function getTotalRatingAvg() {
  var res = APIRequest('issues', {query: [
    {key: 'tracker_id', value: 7},
    {key: 'status_id', value: '*'},
    {key: 'created_on', value: getDateRange()},
    {key: 'cf_7', value: '*'}
  ]});

  var sum = res.issues.reduce(function(a, c) {
    return a + parseInt(c.custom_fields.find(function(i) {return i.id === 7}).value, 10);
  }, 0);

  return res.issues.length ? sum / res.issues.length : 0;
}

function getQuantityRatings(rating) {
  var res = APIRequest('issues', {query: [
    {key: 'tracker_id', value: 7},
    {key: 'status_id', value: '*'},
    {key: 'created_on', value: getDateRange()},
    {key: 'cf_7', value: rating}
  ]});

  return res.issues.length
}

function getWorkUsers(projects) {
  var issues = [];
  projects.forEach(function(project) {
    var res = APIRequest('issues', {query: [
      {key: 'project_id', value: project.id},
      {key: 'tracker_id', value: '!5'},
      {key: 'status_id', value: '*'},
      {key: 'created_on', value: getDateRange()}
    ]});

    issues = issues.concat(res.issues);
  });

  var users = filterUniqueArray(issues.map(function(issue) {return issue.assigned_to;}));

  return users.map(function(user) {
    var userIssues = [];
    projects.forEach(function(project) {
      var res = APIRequest('issues', {query: [
        {key: 'project_id', value: project.id},
        {key: 'assigned_to_id', value: user.id},
        {key: 'tracker_id', value: '!5'},
        {key: 'status_id', value: '*'},
        {key: 'created_on', value: getDateRange()},
        {key: 'cf_7', value: '*'}
      ]});

      userIssues = userIssues.concat(res.issues);
    });

    var sumRatings = userIssues.reduce(function(a, c) {
      return a + parseInt(c.custom_fields.find(function(i) {return i.id === 7}).value, 10);
    }, 0);

    var ratingAvg = userIssues.length ? sumRatings / userIssues.length : 0;

    var timeSpend = 0;
    projects.forEach(function(project) {
      var res = APIRequest('time_entries', {query: [
        {key: 'user_id', value: user.id},
        {key: 'project_id', value: project.id},
        {key: 'spent_on', value: getDateRange()}
      ]});

      timeSpend += res.time_entries.reduce(function(a, c) {
        return a + c.hours;
      }, 0);
    });

    return {
      name: user.name,
      ratingAvg: ratingAvg,
      timeSpend: timeSpend
    }
  });
}

function getTimeSpendByProjects(projects) {
  var timeSpend = 0;
  projects.forEach(function(project) {
    var res = APIRequest('time_entries', {query: [
      {key: 'project_id', value: project.id},
      {key: 'spent_on', value: getDateRange()}
    ]});

    timeSpend += res.time_entries.reduce(function(a, c) {
      return a + c.hours;
    }, 0);
  });

  return timeSpend;
}
