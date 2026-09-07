<!-- BEGIN:nextjs-agent-rules -->
 
# This is NOT the Next.js you know
 
This version has breaking changes — APIs, conventions, and file structure may all differ from your training data. Read the relevant guide in `node_modules/next/dist/docs/` (resolved from this file's directory; in monorepos the `next` package may not be visible from the repo root) before writing any code. Heed deprecation notices.
 
This block is written and re-added by `next dev` — verify at `node_modules/next/dist/server/lib/generate-agent-files.js`. Removing it from a diff only re-creates the uncommitted change; committing it with your work keeps the tree clean.
 
<!-- END:nextjs-agent-rules -->

# Project Guidelines & Agent Rules

## Project Overview
- **Project**: Banchee Web (`banchee-web`)
- **Framework**: Next.js (App Router), React 19
- **Language**: TypeScript (Strict type safety, strictly NO `any`)
- **Styling**: Tailwind CSS
- **UI Components**: shadcn / UI components primarily (Do NOT customize beyond standard patterns; stick to shadcn defaults)
- **State Management**: Zustand
- **Package Manager**: `pnpm` (Strictly NO `npm` or `yarn`)

### Core Goals & Functional Requirements
1. **Document Conversion**:
   - Web-based converter for converting Thai statutory and tax PDF documents into **CSV** and **XLSX** formats.
   - Supported document types:
     - Social Security Fund (ประกันสังคม / SSO)
     - P.N.D. 1 (ภ.ง.ด. 1)
     - P.N.D. 3 (ภ.ง.ด. 3)
     - P.N.D. 53 (ภ.ง.ด. 53)
2. **Custom Document Merging & Form Templating**:
   - Aggregate and merge documents from different companies into customized target form templates according to specific company/client requirements.
3. **UI Localization**:
   - **Primary UI Language: Thai (เน้น UI ภาษาไทย)** for all user-facing interfaces, actions, error handling, notifications, and instructions.
4. **Data Privacy & Stateless Processing (Zero Data Persistence)**:
   - **Strictly DO NOT store or persist any uploaded documents or parsed data**.
   - The application acts solely as an ephemeral converter. All file reading, processing, and file generation must be handled in-memory or discarded immediately after conversion/download. No database or disk persistence of user documents.

---

## Development Commands
- Run development server: `pnpm dev`
- Build for production: `pnpm build`
- Start production server: `pnpm start`

---

## Code Style & Architecture
- **Directory Structure**:
  - `app/`: Next.js App Router pages, layouts, and route handlers.
  - `components/`: Shared React components.
  - `components/ui/`: Reusable UI primitive components (shadcn).
  - `lib/`: Utility functions and shared helpers (e.g., `lib/utils.ts` for `cn`).
  - `stores/`: Zustand stores for client-side state management.
  - `public/`: Static assets.
- **TypeScript**:
  - Maintain strict type safety at all times — **never use `any`** (use `unknown`, generics, or well-defined types/interfaces instead).
  - Always provide explicit interfaces and types for component props, function parameters, state, and data models.
- **Styling & Components**:
  - Use **Tailwind CSS** and **shadcn components** as primary UI building blocks.
  - **Do NOT over-customize** shadcn components or alter their primitive structure unnecessarily; maintain standard shadcn conventions.
  - Merge conditional class names using `cn(...)` from `@/lib/utils`.
  - Ensure all components are modular, accessible, and responsive.
- **State Management**:
  - Use **Zustand** for global and shared client state.
  - Keep stores lean, modular, and separated by feature or domain.

---

## Agent Behavior Guidelines
- Make targeted and surgical modifications; do not refactor unrelated files.
- Preserve existing comments and code formatting where applicable.
- Verify changes using typecheck or build commands when relevant (`pnpm build`).
