DebugExBlackBoard

DebugExBlackBoard is a reference architecture for building structured analytical systems using VBA.

The project demonstrates how Domain-Driven Design, clean architecture principles, and object-oriented modeling can be applied even in environments traditionally dominated by procedural scripting.

Instead of treating spreadsheet columns as raw data fields, this architecture models columns as domain concepts, allowing analytical rules to be implemented as polymorphic domain objects.

Architecture Overview

The system follows a layered architecture inspired by Clean Architecture and Hexagonal Architecture.

Infrastructure
   CSV / IO / Repository
        │
        ▼
Application
   UseCase / Orchestration
        │
        ▼
Domain
   Record → Column → Aggregate → Summary

Key principle:

Domain is the center of the system.
All other layers exist to support the Domain.

Dependency rules:

Application → Domain
Infrastructure → Domain
Application → Infrastructure (via Domain interfaces only)

Reverse dependencies are not allowed.
Domain Model

The domain is designed for analytical processing.

Core concepts:

Record
   Raw domain data

Column
   Domain evaluation rule

ColumnContext
   Evaluation context passed to columns

Aggregate
   Aggregates records and column evaluations

Summary
   Final result produced by aggregates

Processing flow:

Record
   ↓
ColumnContext
   ↓
Column.Evaluate(context)
   ↓
Aggregate
   ↓
Summary
Column Domain Model

A key design decision is treating columns as domain objects rather than simple data fields.

Traditional analytical code often relies on conditional logic such as:

If Grade = 1 And Gender = Male Then

In this architecture, analytical rules are encapsulated inside column objects:

column.Evaluate(context)

This eliminates complex branching and allows the system to evolve through polymorphism.

Column Structure

All columns implement a common interface:

Dom_IEntityColumn
      ▲
  ScalarColumn
  GenderColumn
  GradeColumn
  ClassColumn
  CompositeColumn

Each column represents a domain rule capable of evaluating values from the context.

Composite Columns

Some analytical rules combine multiple columns.

Example:

Grade × Gender
Subject × Grade

This is implemented using the Composite Pattern.

CompositeColumn
   ├ Column
   └ Column

Composite columns allow complex evaluation logic to be built from smaller domain components.

Column Evaluation Context

Column evaluation uses a dedicated context object.

Dom_ColumnContext

Structure:

ColumnContext
   └ Collection<Column>

The context provides the evaluation environment required by column objects.

Design Patterns Used

The architecture intentionally combines several design patterns.

Composite Pattern
Dom_EnrollmentCompositeColumn

Used to compose multiple columns.

Strategy Pattern
Dom_EnrollmentCompositeStrategy

Defines how composite columns combine evaluation results.

Factory Pattern
Dom_EnrollmentColumnFactory
Dom_ClassHourColumnFactory

Responsible for creating column objects.

Dependency Inversion

Repository interfaces are defined in the Domain layer.

Dom_IEnrollmentRepository
Dom_IClassHourRepository

Infrastructure implements these interfaces.

Dependency Rules

The architecture enforces strict dependency direction.

Allowed dependencies:

Application → Domain
Infrastructure → Domain
Application → Infrastructure (via Domain interfaces)

Forbidden dependencies:

Domain → Infrastructure
Domain → IO
Domain → CSV
Domain → FilePath

This ensures the Domain layer remains independent from technical concerns.

Adding a New Column

Adding new analytical rules is straightforward.

Steps:

Create a class implementing Dom_IEntityColumn.

Implement the Evaluate(context) method.

Register the column in the appropriate ColumnFactory.

Example:

Class Dom_SubjectColumn
Implements Dom_IEntityColumn

This approach follows the Open-Closed Principle.

Existing code remains unchanged when new columns are added.

Analytical Domain Modeling

This project models analytical dimensions as domain objects.

Examples:

Grade
Gender
Subject
Class

These behave similarly to dimensions in OLAP systems.

Instead of embedding analytical rules inside procedural logic, they are expressed as domain objects that interact through well-defined interfaces.

Benefits:

• Eliminates complex conditional branching
• Improves extensibility
• Keeps analytical rules inside the Domain layer
• Separates domain logic from infrastructure
Future Extensions

The architecture can be extended to support more advanced analytical scenarios.

Potential extensions:

ColumnDefinition
Column DSL
Expression trees

These would allow dynamic column definitions and expression-based analytical rules.

For the current scope, the existing Column + Composite model provides sufficient flexibility.

Summary

DebugExBlackBoard demonstrates how to implement:

Domain-Driven Design
Hexagonal Architecture
Column Domain Modeling

within a VBA environment.

The project highlights how analytical systems can be structured using:

Domain purity
Polymorphic domain objects
Clear dependency rules
Extensible evaluation models

This repository serves as a reference implementation for building structured analytical systems using VBA.

It demonstrates how analytical logic can be modeled using domain objects, allowing complex evaluation rules to remain extensible, testable, and independent from infrastructure concerns.

The architecture highlights how Domain-Driven Design and clean dependency structures can be applied effectively even in traditional VBA environments.
