VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_RepositoryDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule MoveFieldCloserToUsage, EncapsulatePublicField
'@Folder("Application.DTO")
Option Explicit

Public ScheduleRepo         As Dom_IScheduleRepository
Public SchoolEventRepo      As Dom_ISchoolEventRepository
Public ClassHourRepo        As Dom_IClassHourRepository
Public SubjectRepo          As Dom_ISubjectRepository
Public PeriodRepo           As Dom_IPeriodRepository
Public EnrollmentRepo       As Dom_IEnrollmentRepository
Public MainStreamRepo       As Dom_IMainStreamRepository
Public SpecialStreamRepo    As Dom_ISpecialStreamRepository
