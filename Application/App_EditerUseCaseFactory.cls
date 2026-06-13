VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App_EditerUseCaseFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Application.Factory")
Option Explicit

Private Type Member
    Logger              As App_ILogPersistence
    ScheduleRepo        As Dom_IScheduleRepository
    SchoolEventRepo     As Dom_ISchoolEventRepository
    ClassHourRepo       As Dom_IClassHourRepository
    SubjectRepo         As Dom_ISubjectRepository
    PeriodRepo          As Dom_IPeriodRepository
    EnrollmentRepo      As Dom_IEnrollmentRepository
    MainStreamRepo      As Dom_IMainStreamRepository
    SpecialStreamRepo   As Dom_ISpecialStreamRepository
    Base                As App_BaseUseCase
    SchedulePre         As Pre_DailySchedulePresenter
End Type

Private This As Member

Friend Sub Inject(ByVal Logger As App_ILogPersistence, ByVal ScheduleRepo As Dom_IScheduleRepository, ByVal SchoolEventRepo As Dom_ISchoolEventRepository, ByVal ClassHourRepo As Dom_IClassHourRepository, ByVal SubjectRepo As Dom_ISubjectRepository, ByVal PeriodRepo As Dom_IPeriodRepository, ByVal EnrollmentRepo As Dom_IEnrollmentRepository, ByVal MainStreamRepo As Dom_IMainStreamRepository, ByVal SpecialStreamRepo As Dom_ISpecialStreamRepository, ByVal Base As App_BaseUseCase, ByVal SchedulePre As Pre_DailySchedulePresenter)
    Set This.Logger = Logger
    Set This.ScheduleRepo = ScheduleRepo
    Set This.SchoolEventRepo = SchoolEventRepo
    Set This.ClassHourRepo = ClassHourRepo
    Set This.SubjectRepo = SubjectRepo
    Set This.PeriodRepo = PeriodRepo
    Set This.EnrollmentRepo = EnrollmentRepo
    Set This.MainStreamRepo = MainStreamRepo
    Set This.SpecialStreamRepo = SpecialStreamRepo
    Set This.Base = Base
    Set This.SchedulePre = SchedulePre
End Sub
