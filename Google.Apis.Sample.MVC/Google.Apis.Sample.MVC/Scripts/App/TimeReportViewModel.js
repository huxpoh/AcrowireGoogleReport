var TimeReportViewModel = function (items) {
    var self = this;
    self.items = ko.observableArray(items);
    self.repository = [];

    self.acroItems = ko.observableArray();
    self.acroRepository = [];

    self.acroPassword = ko.observable();
    self.acroLogin = ko.observable();

    self.PopulateAcroArray = function () {
        var login = self.getCookie("login");
        var password = self.getCookie('password');
        if (login == undefined || password == undefined) {
            alert("Set Acro Credentials");
            return;
        }
        else {
            self.acroLogin(login);
            self.acroPassword(password);
        }

        $.ajax({
            url: "/Home/GetAcrowireReport?login="+ login +"&password=" + password,
            async: true
        }).done(function (data) {
            self.acroItems(data.GridList);
            self.acroRepository = data.GridList;
        });
    };

    self.getCookie = function(name) {
        var matches = document.cookie.match(new RegExp(
            "(?:^|; )" + name.replace(/([\.$?*|{}\(\)\[\]\\\/\+^])/g, '\\$1') + "=([^;]*)"
        ));
        return matches ? decodeURIComponent(matches[1]) : undefined;
    };

    self.SetAcroCookie = function() {
        document.cookie = "login=" + self.acroLogin();
        document.cookie = "password=" + self.acroPassword();
        self.PopulateAcroArray();
    };

    self.FilterAcro = function () {
        var result = [];
       
        for (var i = 0; i < self.acroRepository.length; i++) {
            var splitDate = self.acroRepository[i].Date.split('-');
            if (splitDate == "")
                continue;
            var dt = new Date();
            dt.setYear(parseInt(splitDate[0]));
            dt.setMonth(parseInt(splitDate[1])-1);
            dt.setDate(parseInt(splitDate[2]));

            if (self.FilterDateFrom() != "") {
                if (dt < self.FilterDateFrom()) { continue; }
            }

            if (self.FilterDateTo() != "") {
                if (dt > self.FilterDateTo()) { continue; }
            }

            if (self.FilterProject() != "") {
                if (self.acroRepository[i].Project != self.FilterProject()) { continue; }
            }

            result.push(self.acroRepository[i]);
        }
        self.acroItems(result);
    };


    self.FilterProject = ko.observable("");
    self.FilterTotalHours = ko.observable(0);
    self.FilterTotalOverTime = ko.observable(0);

    self.FilterDateFrom = ko.observable("");
    self.FilterDateTo = ko.observable("");

    self.PopulateArray = function () {
        $.ajax({
            url: "/Home/TimeReport",
            async: true,
            context: document.body
        }).done(function (data) {
            self.items(data.GridList);
            self.repository = data.GridList;
        });
    };

    self.Filter = function () {
        var result = [];
        self.FilterTotalHours(0);
        for (var i = 0; i < self.repository.length; i++) {
            var splitDate = self.repository[i].Date.split('-');
            if (splitDate == "")
                continue;
            var dt = new Date();
            dt.setYear(parseInt(splitDate[0]));
            dt.setMonth(parseInt(splitDate[1])-1);
            dt.setDate(parseInt(splitDate[2]));

            if (self.FilterDateFrom() != "") {
                if (dt < self.FilterDateFrom()) {
                    continue;
                }
            }

            if (self.FilterDateTo() != "") {
                if (dt > self.FilterDateTo()) {
                     continue;
                }
            }

            if (self.FilterProject() != "") {
                if (self.repository[i].Project != self.FilterProject()) {
                     continue;
                }
            }

            self.FilterTotalHours(self.FilterTotalHours() + self.repository[i].Duration);
            self.FilterTotalOverTime(self.FilterTotalOverTime() + self.repository[i].Overtime);
            result.push(self.repository[i]);
        }
        self.items(result);
    };

    self.ResetFilter = function () {
        self.FilterTotalHours(0);
        self.items(self.repository);
        self.acroItems(self.acroRepository);
        self.FilterProject("");
    };

    self.PopulateArray();
    self.PopulateAcroArray();
};