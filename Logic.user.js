// ==UserScript==
// @name         Logic
// @namespace    http://tampermonkey.net/
// @version      0.1
// @description  try to take over the world!
// @author       You
// @match        https://tasks.office.com/FR/Home/Planner
// @require      https://code.jquery.com/jquery-3.5.1.min.js
// @grant        none
// ==/UserScript==

class API {
    constructor() {
        this.base_url = {
            'task' : 'https://tasks.office.com/jwsite.onmicrosoft.com/TasksApiV1/',
            'graph' : 'https://tasks.office.com/jwsite.onmicrosoft.com/GraphApiV1/'
        };
    }

    constructURL(base, endpoint, options) {
        let url = this.base_url[base] + encodeURI(endpoint);

        if (options) {
            url += '?';
        }

        for (let key in options) {
            url += encodeURI(key);
            url += '=';
            url += encodeURI(options[key]);
        }
        return url;
    }


    async get(url) {
        let response = await fetch(url, {method: 'GET', 'headers': window.HEADERS});
        let data = await response.json();
        return data;
    }

    async put(url, data) {
        let response = await fetch(url, {method: 'PUT', body: JSON.stringify(data), 'headers': window.HEADERS});
        let data_resp = await response.json();
        return data_resp;
    }

    async post(url, data) {
        let response = await fetch(url, {method: 'POST', body: JSON.stringify(data), 'headers': window.HEADERS});
        let data_resp = await response.json();
        return data_resp;
    }

    async getTasksForPlan(id) {
        let url = this.constructURL('task', 'GetTasksForPlan', { 'planId' : id });
        return (await this.get(url)).Results;
    }

    async getPlan(id) {
        let url = this.constructURL('task', 'GetPlan', { 'planId' : id });
        return await this.get(url);
    }

    async getBucketsInPlan(id) {
        let url = this.constructURL('task', 'GetBucketsInPlan', { 'planId' : id });
        return (await this.get(url)).Results.map(value => value.Bucket);
    }

    async getGroups() {
        let url = this.constructURL('graph', 'GetAllGroupsForCurrentUserAsync', {'skipToken': ''});
        let data = await this.get(url);
        let groups = [];
        for (let group of data.groups) {
            groups.push({name: group.DisplayName, id: group.Id});
        }

        return groups;
    }

    async getTasksDetail(tasks) {
        let url = this.constructURL('task', 'GetTaskDetailsBatchedAsync');
        let request = {
            taskIds: tasks.map(task => task.Task.Id)
        };
        let data = await this.put(url, request);
        return data;
    }

    async getPlans(groups) {
        let url = this.constructURL('task', 'ResolveGroupsToPlansBatchedAsync');
        let request = {
            groupIds: groups.map(group => group.id)
        }
        let data = await this.put(url, request);
        let plans = [];
        for (let group in data) {
            data[group].forEach(plan => plans.push({title: plan.Title, id: plan.Id, group_id: group}));
        }

        return plans;
    }

    async createBucket(plan_id, bucket_name) {
        let url = this.constructURL('task', 'CreateBucket');
        let request = {
            Bucket: {
             Title: bucket_name,
             PlanId: plan_id,
            }
        };
        return await this.post(url, request);
    }

  async createTask(task, details) {
   let url = this.constructURL('task', 'CreateTask');
   let request = {
      newTask: task,
      newTaskDetails: JSON.stringify(details),
      newTaskFormattings: {
          BucketBoardFormat: {
              //Id: 'PLEXTID263',
              ItemVersion: null,
              //OrderHint: " !",
              TaskBoardType: 2
          }
      }
   }
   return await this.post(url, request);
  }

}

const App = {
  api : new API(),
  current_plan: new URLSearchParams(window.location.href).get('planId'),
  groups: [],
  plans: [],
  buckets: [],
  tasks: [],
  filtered_plans: [],
  target : {
   id : ''
  }
}

const api = new API();
var plan_id


function updatePlans() {
    let cgroup = $('#groups').val();
    App.filtered_plans = App.plans.filter(plan => plan.group_id == cgroup);
    $('#plans').children().remove().end()
    App.filtered_plans.forEach(plan => {
        let opt = $('<option/>');
        opt.text(plan.title).attr('value', plan.id);
        $('#plans').append(opt);
    });
}

function updateGroups() {
    $('#groups').children().remove().end()
    App.groups.forEach(group => {
        let opt = $('<option/>');
        opt.text(group.name).attr('value', group.id);
        $('#groups').append(opt);
    });
}

function updateTarget() {
    App.target.id = $('#plans').val();
}


async function copyBucket(bucket, target_plan) {
    let new_bucket = await App.api.createBucket(target_plan, bucket.Title);
    let bucket_tasks = App.tasks.filter(task => task.Task.BucketId == bucket.Id);
    bucket_tasks.forEach(task => {
        copyTask(task, new_bucket.Bucket.Id, target_plan);
    });

}

async function copyTask(task, target_bucket, target_plan) {
    let detail = await App.api.getTasksDetail([task]);
    detail = detail[task.Task.Id];
    //detail.Id = 'PLEXTID263';
    delete detail.Id;
    delete detail.Type;
    delete detail.CompletedBy;
    detail.ItemVersion = 0;



    let hints = Object.keys(detail.Checklist).map(key => detail.Checklist[key].OrderHint);
    hints.sort(ordinalSort);
    hints[-1] = "";

    console.log(hints)

    for (let [key, checklist] of Object.entries(detail.Checklist)) {
        console.log(key)
        console.log(checklist)
        let index = hints.indexOf(checklist.OrderHint);
        checklist.OrderHint = hints[index] + " " + hints[index-1] + "!";
    }

    console.log(detail);

    let t = task.Task;
    //t.AppliedCategories = []; //todo voir si possible (etiquette)
    t.Assignments = {};
    delete t.CompletedBy;
    delete t.CompletedDate;
    delete t.ConversationThreadId;
    t.ItemVersion= null;
    //t.OrderHint = " !";
    t.OrderHint = "";
    //t.Id = 'PLEXTID263';
    delete t.Id;
    t.PlanId = target_plan;
    t.BucketId = target_bucket;
    delete t.Type;

    let new_task = await App.api.createTask(t, detail);
}

function ordinalSort(a, b) {
    return (a == b ? 0 : (a > b ? -1 : 1));
}

async function launchCopy() {
    App.buckets.forEach(bucket => {
        copyBucket(bucket, App.target.id );
    });
}

setTimeout(async () => {
    App.groups = await App.api.getGroups();
    App.plans = await App.api.getPlans(App.groups);

    $('#groups').on('change', () => updatePlans());
    $('#plans').on('change', () => updateTarget());
    updateGroups();


    App.buckets = await App.api.getBucketsInPlan(App.current_plan);
    App.tasks = await App.api.getTasksForPlan(App.current_plan);
    $('#copy').on('click', () => launchCopy());

  /*$('#copy').click(e => {
     let buckets = [];
     let tasks = [];

    App.api.getBucketsInPlan(App.current_plan).then((data) => {
        data.Results.forEach(bucket => buckets.push({ title: bucket.Bucket.Title, id: bucket.Bucket.Id }));
        console.log(buckets);
    });


  });*/
}, 4000);




