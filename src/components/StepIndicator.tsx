interface Step {
  number: number;
  label: string;
  description?: string;
}

interface StepIndicatorProps {
  steps: Step[];
  currentStep: number;
}

export function StepIndicator({ steps, currentStep }: StepIndicatorProps) {
  return (
    <ol className="langkah" aria-label="Langkah kerja">
      {steps.map((step) => {
        const isCompleted = step.number < currentStep;
        const isCurrent = step.number === currentStep;
        return (
          <li
            key={step.number}
            className={isCurrent ? 'aktif' : isCompleted ? 'selesai' : undefined}
            aria-current={isCurrent ? 'step' : undefined}
          >
            <span className="angka">{isCompleted ? '✓' : step.number}</span>
            {step.label}
            {step.description && <small>{step.description}</small>}
          </li>
        );
      })}
    </ol>
  );
}
