import { Check } from 'lucide-react';

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
    <div className="flex items-start gap-0">
      {steps.map((step, idx) => {
        const isCompleted = step.number < currentStep;
        const isCurrent = step.number === currentStep;
        const isFuture = step.number > currentStep;

        return (
          <div key={step.number} className="flex-1 flex items-center">
            {/* Step circle + label */}
            <div className="flex flex-col items-center">
              <div
                className={[
                  'w-8 h-8 rounded-full flex items-center justify-center text-xs font-bold border-2 transition-all duration-300',
                  isCompleted
                    ? 'bg-emerald-500 border-emerald-500 text-white'
                    : isCurrent
                      ? 'bg-blue-600 border-blue-600 text-white shadow-md shadow-blue-200 scale-110'
                      : 'bg-white border-gray-200 text-gray-400',
                ].join(' ')}
              >
                {isCompleted ? (
                  <Check className="w-4 h-4" />
                ) : (
                  step.number
                )}
              </div>
              <div className="mt-1.5 text-center">
                <span
                  className={`text-xs font-semibold ${
                    isCurrent ? 'text-blue-600' : isCompleted ? 'text-emerald-600' : 'text-gray-400'
                  }`}
                >
                  {step.label}
                </span>
                {step.description && (
                  <p className={`text-[10px] mt-0.5 ${isCurrent ? 'text-blue-500' : 'text-gray-400'}`}>
                    {step.description}
                  </p>
                )}
              </div>
            </div>

            {/* Connector line */}
            {idx < steps.length - 1 && (
              <div className="flex-1 mx-2 mt-4">
                <div
                  className={`h-0.5 rounded-full transition-all duration-500 ${
                    isCompleted ? 'bg-emerald-300' : 'bg-gray-200'
                  }`}
                />
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
