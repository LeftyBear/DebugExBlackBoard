DebugExBlackBoard
Overview

This project follows a Domain-Centered Architecture.

The design prioritizes:

Strict layer responsibility separation

Dependency Inversion

Structural clarity over convenience

Reduction of unnecessary classes

Deterministic review standards

This repository is intentionally structured for architectural review.

Architecture Principle
Core Philosophy

The Domain layer is the center of the system.

All other layers exist to support the Domain.

Dependency direction must always be:

Application → Domain
Infrastructure → Domain
Application → Infrastructure (only via Domain interface)

Reverse dependencies are strictly forbidden.

Layer Definitions
Domain Layer

Responsibility: Meaning and Business Rules

Holds business concepts and invariants

Defines repository interfaces

Contains no IO logic

Contains no file paths

Contains no external system knowledge

Must not reference Infrastructure

Naming:
Prefix: Dom_

Application Layer

Responsibility: Use Case Orchestration

Coordinates Domain objects

Controls execution flow

Owns no business rule logic

Owns no mapping/dictionary

Owns no file path

Rules:

All arguments must be explicitly ByVal

Selector and Resolver must be separated

No mapping logic inside Application

Must not implement Infrastructure interfaces

Naming:
Prefix: App_

Infrastructure Layer

Responsibility: External Interaction

File IO

CSV reading

Database access

External APIs

Rules:

Implements Domain-defined interfaces

Must not depend on Application layer

Must not contain business rules

Purely technical implementation

Naming:
Prefix: Inf_

Utility Layer

Responsibility: Stateless Pure Functions

No state retention

No side effects

No layer dependency

Pure transformation only

Naming:
Prefix: Util_

Coding Canon

This project follows strict coding conventions.

Naming Rules

PascalCase required

Explicit prefix per layer

No ambiguous naming

VBA Structural Rules

Use Private Type Member

Use Private This As Member

Access fields via This.FieldName

No line continuation character "_"

No implicit ByRef

Always declare ByVal explicitly

Dependency Rules

Domain must not reference Infrastructure

Infrastructure must not reference Application

No circular dependencies

Dependency Inversion required

Forbidden Patterns

The following are strictly prohibited:

Domain referencing Infrastructure

Application owning file paths

Application containing mapping dictionaries

Infrastructure containing business logic

Selector and Resolver combined

Hidden state in Utility

Layer mixing inside single class

Any violation must be flagged during review.

Review Mode Specification

When reviewing this repository:

Prioritize architecture violations

Detect dependency direction errors

Identify responsibility leakage

Suggest class reduction opportunities

Reject convenience-driven shortcuts

Prefer structural correctness over brevity

Architecture integrity must be evaluated before implementation detail.

Design Intent

This repository exists to:

Validate Domain-centered design in VBA

Establish a reusable canonical structure

Maintain deterministic architectural integrity

Enable consistent long-term evolution

Usage for Review

To review this repository, state:

Review this repository based on its README architecture canon.

The README must be treated as the authoritative architectural source.

Final Principle

Structural integrity is more important than implementation speed.
