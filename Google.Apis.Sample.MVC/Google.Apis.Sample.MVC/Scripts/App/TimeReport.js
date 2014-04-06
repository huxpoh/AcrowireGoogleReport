function timeReport() {
    var self = this;
    self.Date = ko.observable("");
    self.Project = ko.observable("");
    self.Task = ko.observable("");
    self.Type = ko.observable("");
    self.Description = ko.observable("");
    self.Duration = ko.observable(0);
    self.Overtime = ko.observable(0);
};